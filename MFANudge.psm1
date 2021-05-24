#requires -Version 5.0

Set-StrictMode -Version 'Latest'

$MSALModuleName = 'msal.ps'

$ConfigFileName = '.\mfanudgeconfig.json'
$MyConfig = $null


$MSALToken = $null
$Headers   = $null

$MSALModuleName = 'msal.ps'

$RedirectURI = 'https://login.microsoftonline.com/common/oauth2/nativeclient'
$NudgePolicyURI =  'https://graph.microsoft.com/beta/policies/authenticationmethodspolicy'



# Output messages
$CONFIGFILENOTFOUND     = 'Configuration information not found. Please run Save-NudgeModuleConfig to setup the module configuration.'
$INVALIDTENANTID        = 'TenantID is not a valid GUID'
$INVALIDCLIENTID        = 'ClientID is not a valid GUID'
$CANNOTSAVECONFIG       = 'Unable to save configuration file'
$CONFIGFILESAVED        = 'Configuration has been saved'
$MSALNOTINSTALLED       = 'Please install msal.ps before using this module'
$UNABLETOGETACCESSTOKEN = 'Unable to get an access token'
$DISABLEPOLICYFAILED    = 'Unable to disable the nudge policy'
$DISABLEPOLICYOK        = 'Nudge Policy disabled'

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

 #   try 
 #   {
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
  #  }
  #  catch
  #  {
  #      throw $UNABLETOGETACCESSTOKEN
  #  }

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

    if ($Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
    {
        Write-Host $DISABLEPOLICYFAILED
    }
    else
    {
        Write-Host $DISABLEPOLICYOK
    }
}
Export-ModuleMember -Function Disable-MFANudge

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