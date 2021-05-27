#requires -Version 5.0

Set-StrictMode -Version 'Latest'

# MSAL and token variables
$MSALModuleName = 'msal.ps'
$MSALToken = $null
$Headers   = $null

# Configuration file variables
$ConfigFileName = '.\mfanudgeconfig.json'
$MyConfig = $null

# MSGraph variables
$RedirectURI    = 'https://login.microsoftonline.com/common/oauth2/nativeclient'
$NudgePolicyURI = 'https://graph.microsoft.com/beta/policies/authenticationmethodspolicy'
$UsersURI       = 'https://graph.microsoft.com/v1.0/users'
$GroupsURI      = 'https://graph.microsoft.com/v1.0/groups'

# Output messages
$CONFIGFILENOTFOUND       = 'Configuration information not found. Please run Save-NudgeModuleConfig to setup the TenantID and CLientID configuration.'
$INVALIDTENANTID          = 'TenantID is not a valid GUID'
$INVALIDCLIENTID          = 'ClientID is not a valid GUID'
$CANNOTSAVECONFIG         = 'Unable to save configuration file'
$CONFIGFILESAVED          = 'Configuration has been saved'
$MSALNOTINSTALLED         = 'Please install msal.ps before using this module'
$UNABLETOGETACCESSTOKEN   = 'Unable to get an access token'
$DISABLEPOLICYFAILED      = 'Unable to disable the nudge policy'
$DISABLEPOLICYOK          = 'Nudge Policy disabled'
$GETPOLICYFAILED          = 'Failed to get Nudge policy'
$ENABLEPOLICYFAILED       = 'Enable policy failed'
$ENABLEPOLICYOK           = 'Enable policy succeeded'
$INCLUDELISTEMPTY         = 'Include list is empty'
$ALLEXCLUDETARGETSIGNORED = 'All the exclude targets were ignored. Policy will not be modified'
$POLICYNOTFOUND           = 'Nudge policy not found'
$REGENFENTRYNOTFOUND      = 'Registration enforcement entry not found'
$REGCAMNOTFOUND           = 'Registration campaign entry not found'
$INCLUDEPARAMSEMPTY       = 'At least one entry should be included with -IncludeUsers or -IncludeGroups'

function Import-NudgeModuleConfig
{
    if (!(Test-Path -Path $ConfigFileName -PathType Leaf))
    {
        return $false
    }

    try 
    {
        $Script:MyConfig = Get-Content -Path $ConfigFileName -Encoding utf8 -Raw | ConvertFrom-Json
    }
    catch 
    {
        return $false
    }

    Write-Debug $Script:MyConfig

    return $true
}

function Save-NudgeModuleConfig
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string]
        $TenantID,

        [Parameter(Mandatory=$True)]
        [string]
        $ClientID
    )

    $TenantIDGuid = [guid]::Empty
    $ClientIDGuid = [guid]::Empty

    if (![guid]::TryParse($TenantID, [ref]$TenantIDGuid))
    {
        Write-Error $INVALIDTENANTID
        return $false
    }


    if (![guid]::TryParse($ClientID, [ref]$ClientIDGuid))
    {
        Write-Error $INVALIDCLIENTID
        return $false
    }

    $tempConfig = [pscustomobject]@{TenantID = $TenantID; ClientID = $ClientID}
    try 
    {
        $tempConfig | ConvertTo-Json | Out-File -FilePath $ConfigFileName -Encoding utf8 -Force        
    }
    catch
    {
        Write-Host $CANNOTSAVECONFIG
        return $false    
    }

    $Script:MyConfig = $tempConfig

    Write-Host $CONFIGFILESAVED
#    return $True
}
Export-ModuleMember -Function 'Save-NudgeModuleConfig'

function Get-Token
{
    param ()
    
    $Result = $null

    try 
    {
        if (!$script:MSALToken)
        {
            $Result = Get-MsalToken -TenantId $MyConfig.TenantID -ClientId $MyConfig.ClientID -RedirectUri $RedirectURI -Interactive
        }
        else
        {
            $Result = Get-MsalToken -TenantId $MyConfig.TenantID -ClientId $MyConfig.ClientID -RedirectUri $RedirectURI -Silent
            if (!$Result)
            {
                $Result = Get-MsalToken -TenantId $MyConfig.TenantID -ClientId $MyConfig.ClientID -Interactive
            }       
        }            
    }
    catch
    {
       throw $UNABLETOGETACCESSTOKEN
    }

    $script:MSALToken = $Result
    $Script:Headers = @{"Authorization" = "Bearer $($Result.AccessToken)"; "Content-Type" = "application/json"}

    return $true
}

function Invoke-Graph
{
    [CmdletBinding()]

    param
    (
        [Parameter(Mandatory=$True)]
        [string]
        $URI,

        [Parameter(Mandatory=$True)]
        [string]
        $Method,

        [Parameter(Mandatory=$False)]
        [string]
        $Body
    )
      
    do 
    {
        try
        {
            if ($Body)
            {
                $Result = Invoke-WebRequest -Uri $URI -Method $Method -Body $Body -Headers $Script:Headers -UseBasicParsing
            }
            else
            {
                $Result = Invoke-WebRequest -Uri $URI -Method $Method -Headers $Script:Headers -UseBasicParsing
            }
            
        }
        catch [System.Net.WebException]
        {
            if (([int]$_.Exception.Response.StatusCode) -eq [System.Net.HttpStatusCode]::Unauthorized)
            {
                if (Get-Token)
                {
                    continue
                }
            }
            return [pscustomobject]@{StatusCode = [System.Net.HttpStatusCode]($_.Exception.Response.StatusCode); Content = $null}
        }
        catch
        {
            return [pscustomobject]@{StatusCode = $null; Content = $null}
        }
        break
    } while ($True)

    return [pscustomobject]@{StatusCode = [System.Net.HttpStatusCode]($Result.StatusCode); Content = $Result.Content}
}

function Disable-MFANudge
{
    [CmdletBinding()]
    
    param()

    if (!$MyConfig)
    {
        Write-Host $CONFIGFILENOTFOUND
        return
    }

    $DisableNudgeJSON = '{"registrationEnforcement": { "authenticationMethodsRegistrationCampaign": { "state": "disabled" } } }'

    try 
    {
        $Result = Invoke-Graph -URI $NudgePolicyURI -Method 'PATCH' -Body $DisableNudgeJSON
    }
    catch 
    {
        Write-Host $DISABLEPOLICYFAILED
        return
    }

    if ($Result.StatusCode -ne [System.Net.HttpStatusCode]::NoContent)
    {
        Write-Host $DISABLEPOLICYFAILED
    }
    else
    {
        Write-Host $DISABLEPOLICYOK
    }
}
Export-ModuleMember -Function Disable-MFANudge

function Convert-UserUPNToGUID
{
    param
    (
        [String]
        $UPN
    )

    if ($UPN.Length -eq 0)
    {
        return [GUID]::Empty
    }

    $UserQueryURI = $UsersURI + '/' + [System.Web.HttpUtility]::UrlEncode($UPN)

    $Result = Invoke-Graph -URI $UserQueryURI -Method 'GET'

    if (($Result.StatusCode -ne [System.Net.HttpStatusCode]::OK) -or ($Result.Content.Length -eq 0))
    {
        return [GUID]::Empty
    }

    return [GUID](($Result.Content | ConvertFrom-Json).id)
}

function Convert-UserGUIDToUPN
{
    [CmdletBinding()]

    param
    (
        [String]
        $GUID
    )

    if ($GUID.Length -eq 0)
    {
        return ''
    }

    $UserQueryURI = $UsersURI + '/' + [System.Web.HttpUtility]::UrlEncode($GUID)
    $Result = Invoke-Graph -URI $UserQueryURI -Method 'GET'

    if (($Result.StatusCode -ne [System.Net.HttpStatusCode]::OK) -or ($Result.Content.Length -eq 0))
    {
        return ''
    }

    return (($Result.Content | ConvertFrom-Json).userPrincipalName)
}

function Convert-GroupNameToGUID
{
    param
    (
        [String]
        $GroupName
    )

    if ($GroupName.Length -eq 0)
    {
        return [GUID]::Empty
    }

    $GroupQueryURI = $GroupsURI + '?$filter=(displayName eq ''' + $GroupName + ''')'

    $Result = Invoke-Graph -URI $GroupQueryURI -Method 'GET'

    if (($Result.StatusCode -ne [System.Net.HttpStatusCode]::OK) -or ($Result.Content.Length -eq 0))
    {
        return [GUID]::Empty
    }
    
    $Obj = ($Result.Content | ConvertFrom-Json).value
    if ($Obj.Count -ne 1)
    {
        return [GUID]::Empty
    }
    return $Obj[0].id
}

function Convert-GroupGUIDToName
{
    [CmdletBinding()]

    param
    (
        [String]
        $GUID
    )

    if ($GUID.Length -eq 0)
    {
        return ''
    }

    $GroupQueryURI = $GroupsURI + '/' + [System.Web.HttpUtility]::UrlEncode($GUID)
    $Result = Invoke-Graph -URI $GroupQueryURI -Method 'GET'

    if (($Result.StatusCode -ne [System.Net.HttpStatusCode]::OK) -or ($Result.Content.Length -eq 0))
    {
        return ''
    }

    return (($Result.Content | ConvertFrom-Json).displayName)
}

function Get-MSGraphNudgePolicy
{
    param()

    try 
    {
        $Result = Invoke-Graph -URI $NudgePolicyURI -Method 'GET'
    }
    catch 
    {
        return $null
    }

    if ($Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
    {
        Write-Host $GETPOLICYFAILED
        return $null
    }

    return ($Result.Content | ConvertFrom-Json)
}

function Get-MFANudge
{
    [CmdletBinding()]
    
    param()

    if (!$MyConfig)
    {
        Write-Host $CONFIGFILENOTFOUND
        return
    }

    $Policy = Get-MSGraphNudgePolicy

    if (!$Policy)
    {
        Write-Host $POLICYNOTFOUND
        return
    }

    if ($Policy.psobject.Properties.Name -notcontains 'registrationEnforcement')
    {
        Write-Host $REGENFENTRYNOTFOUND
        return
    }

    $RegEnf = $Policy.registrationEnforcement
    if ($RegEnf.psobject.Properties.Name -notcontains 'authenticationMethodsRegistrationCampaign')
    {
        Write-Host $REGCAMNOTFOUND
        return
    }

    $Nudge = $RegEnf.authenticationMethodsRegistrationCampaign

    Write-Host "Nudge policy is $($Nudge.state)"

    # No more info to show if policy is disabled
    if ($Nudge.state -eq 'disabled')
    {
        return
    }

    Write-host "Snooze duration is $($Nudge.snoozeDurationInDays) days"

    Write-Host ''
    Write-Host 'Included targets:'
    foreach ($i in $Nudge.includeTargets)
    {
        if ($i.targetType -eq 'user')
        {
            $DisplayName = 'User: ' + (Convert-UserGUIDToUPN -GUID ($i.id))
        }
        else
        {
            $DisplayName = 'Group: '
            if ($i.id -ne 'All_Users')
            {
                $DisplayName += Convert-GroupGUIDToName -GUID $i.id
            }
            else
            {
                $DisplayName += 'All_Users'
            }
        }

        $DisplayName += " (GUID: $($i.id))"
        Write-Host $DisplayName
    }

    Write-Host ''
    Write-Host 'Excluded targets:'
    if ($Nudge.psobject.Properties.name -contains 'excludeTargets' -and $Nudge.excludeTargets.Count -gt 0)
    {
        foreach ($i in $Nudge.excludeTargets)
        {
            if ($i.targetType -eq 'user')
            {
                $DisplayName = 'User: ' + (Convert-UserGUIDToUPN -GUID ($i.id))
            }
            else
            {
                $DisplayName = 'Group: ' + (Convert-GroupGUIDToName -GUID ($i.id))
            }
    
            $DisplayName += " (GUID: $($i.id))"
            Write-Host $DisplayName
        }        
    }
    else
    {
        Write-Host 'No excluded targets'
    }
    Write-Host ''
}
Export-ModuleMember -Function Get-MFANudge

function Enable-MFANudge
{
    [CmdletBinding()]
    
    param
    (
        [Parameter()]
        [ValidateRange(0, 14)]
        [UInt16]
        $SnoozeDuration = 0,

        [Parameter(ParameterSetName='AllUsers')]
        [Switch]
        $IncludeAllUsers,

        [Parameter(ParameterSetName='ScopedInclude')]
        [String[]]
        $IncludeUsers,

        [Parameter(ParameterSetName='ScopedInclude')]
        [String[]]
        $IncludeGroups,

        [Parameter()]
        [String[]]
        $ExcludeUsers,

        [Parameter()]
        [String[]]
        $ExcludeGroups
    )

    if (!$MyConfig)
    {
        Write-Host $CONFIGFILENOTFOUND
        return
    }

    $IncludeTargets = [System.Collections.ArrayList]::new()

    if ($IncludeAllUsers)
    {
        $AllUsersEntry = [pscustomobject]@{targetType='group'; id='All_users'; targetedAuthenticationMethod='microsoftAuthenticator'}
        $IncludeTargets = @($AllUsersEntry)
    }
    else
    {
        if ((($null -eq $IncludeUsers) -or ($IncludeUsers.Count -eq 0)) -and (($null -eq $IncludeGroups) -or ($IncludeGroups.Count -eq 0)))
        {
            Write-Host $INCLUDEPARAMSEMPTY
            return
        }

        foreach ($u in $IncludeUsers)
        {
            $GUID = Convert-UserUPNToGUID -UPN $u
            if ($GUID -eq [GUID]::Empty)
            {
                Write-Host "Warning: cannot find user with UPN $u. Entry will be ignored"
                continue
            }

            [void]$IncludeTargets.Add([pscustomobject]@{targetType='user'; id=$GUID.ToString(); targetedAuthenticationMethod='microsoftAuthenticator'})
        }

        foreach ($g in $IncludeGroups)
        {
            $GUID = Convert-GroupNameToGUID -GroupName $g
            if ($GUID -eq [GUID]::Empty)
            {
                Write-Host "Warning: cannot find group with name $g. Entry will be ignored"
                continue
            }

            [void]$IncludeTargets.Add([pscustomobject]@{targetType='group'; id=$GUID.ToString(); targetedAuthenticationMethod='microsoftAuthenticator'})
        }

        if ($IncludeTargets.Count -eq 0)
        {
            Write-Host $INCLUDELISTEMPTY
            return
        }
    }

    $ExcludeTargets = [System.Collections.ArrayList]::new()
    foreach ($u in $ExcludeUsers)
    {
        # Map the user to GUID
        if ($u -in $IncludeUsers)
        {
            Write-Host "User $u is in the include and exclude list. Ignorting exclude entry"
            continue
        }

        $GUID = Convert-UserUPNToGUID -UPN $u
        if ($GUID -eq [GUID]::Empty)
        {
            Write-Host "Warning: cannot find user with UPN $u. Entry will be ignored"
            continue
        }

        [void]$ExcludeTargets.Add([pscustomobject]@{targetType='user'; id=$GUID.ToString(); targetedAuthenticationMethod='microsoftAuthenticator'})
    }

    foreach ($g in $ExcludeGroups)
    {
        if ($g -in $IncludeGroups)
        {
            Write-Host "Group $g is in the include and exclude list. Ignorting exclude entry"
            continue
        }   
        
        $GUID = Convert-GroupNameToGUID -GroupName $g
        if ($GUID -eq [GUID]::Empty)
        {
            Write-Host "Warning: cannot find group with name $g. Entry will be ignored"
            continue
        }

        [void]$ExcludeTargets.Add([pscustomobject]@{targetType='group'; id=$GUID.ToString(); targetedAuthenticationMethod='microsoftAuthenticator'})        
    }

    # Do not set the policy if all the entries in the ExcludeUsers and ExcludeGroups were ignored
    if ((($ExcludeUsers.Count -gt 0) -or ($ExcludeGroups.Count -gt 0)) -and ($ExcludeTargets.Count -eq 0))
    {
        Write-Host $ALLEXCLUDETARGETSIGNORED
        return
    }
    
    $RegEnfJSON = [pscustomobject]@{registrationEnforcement=[pscustomobject]@{authenticationMethodsRegistrationCampaign=[pscustomobject]@{state='enabled'; snoozeDurationInDays=$SnoozeDuration; includeTargets=$IncludeTargets; excludeTargets=$ExcludeTargets} } } | ConvertTo-Json -Compress -Depth 99

    try 
    {
        $Result = Invoke-Graph -URI $NudgePolicyURI -Method 'PATCH' -Body $RegEnfJSON
    }
    catch 
    {
        Write-Host $ENABLEPOLICYFAILED
        return
    }

    if ($Result.StatusCode -ne [System.Net.HttpStatusCode]::NoContent)
    {
        Write-Host $ENABLEPOLICYFAILED
    }
    else
    {
        Write-Host $ENABLEPOLICYOK
    }
}
Export-ModuleMember -Function Enable-MFANudge

function Set-MFANudgeSnoozeDuration
{
    [CmdletBinding()]
    
    param
    (
        [Parameter(Position=0)]
        [ValidateRange(0, 14)]
        [UInt16]
        $SnoozeDuration = 0
    )

    if (!$MyConfig)
    {
        Write-Host $CONFIGFILENOTFOUND
        return
    }

    $Policy = Get-MSGraphNudgePolicy

    if (!$Policy)
    {
        Write-Host 'Nudge policy not found'
        return
    }

    if ($Policy.psobject.Properties.Name -notcontains 'registrationEnforcement')
    {
        Write-Host $REGENFENTRYNOTFOUND
        return
    }

    if ($Policy.registrationEnforcement.psobject.Properties.Name -notcontains 'authenticationMethodsRegistrationCampaign')
    {
        Write-Host $REGCAMNOTFOUND
        return
    }

    if ($Policy.registrationEnforcement.authenticationMethodsRegistrationCampaign.state -eq 'Disabled')
    {
        Write-Host 'Nudge policy is disabled. Enable the policy using Enable-MFANudge'
        return
    }

    if ($Policy.registrationEnforcement.authenticationMethodsRegistrationCampaign.snoozeDurationInDays -eq $SnoozeDuration)
    {
        Write-Host "Snooze duration is already set to the value $SnoozeDuration days. No further changes required."
        return
    }

    $Policy.registrationEnforcement.authenticationMethodsRegistrationCampaign.snoozeDurationInDays = $SnoozeDuration
    $JSON = $Policy | ConvertTo-Json -Compress -Depth 99

    try 
    {
        $Result = Invoke-Graph -URI $NudgePolicyURI -Method 'PATCH' -Body $JSON
    }
    catch 
    {
        Write-Host 'Failed to set snooze'
        return
    }

    if ($Result.StatusCode -ne [System.Net.HttpStatusCode]::NoContent)
    {
        Write-Host 'Failed to set snooze'
    }
    else
    {
        Write-Host "Snooze duration set to $SnoozeDuration"
    }
}
Export-ModuleMember -Function Set-MFANudgeSnoozeDuration

# Main
if (!(Import-NudgeModuleConfig))
{
    Write-Host $CONFIGFILENOTFOUND
}

if (-not (Get-Module -Name $MSALModuleName -ListAvailable))
{
    Write-Host $MSALNOTINSTALLED
    throw
}