#requires -modules "PnP.PowerShell"

[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol   = [System.Net.SecurityProtocolType]::Tls12

$clientId   = $env:O365_CLIENTID
$thumbprint = $env:O365_THUMBPRINT
$tenantId   = $env:O365_TENANTID
$tenant     = $env:O365_TENANT

# you will see this as output in the Deploy-AzureResources.ps1 script or pull directly from the Logic App's trigger action.
$powerAutomateOrLogicAppTriggerUrl = "https://prod-50.eastus.logic.azure.com:443/workflows/116677db9ed04621b15a49e04b3c6da9/triggers/manual/paths/invoke?api-version=2019-05-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=GImmL7u6oApxOC0XsMa_A_bwHzbfDyUMxEITxMDpxj8"

# manually upload these files and update the paths below
$previewImageUrl  = "https://$tenant.sharepoint.com/SiteAssets/contoso-logo.jpeg"


# connect to tenant admin

    Connect-PnPOnline `
        -Url        "https://$tenant-admin.sharepoint.com" `
        -ClientId   $clientId `
        -Thumbprint $thumbprint `
        -Tenant     $tenantId


# create site script 

    $template = '
    {{
    "$schema" : "schema.json",
    "actions" : [
        {{
        "verb" : "triggerFlow",
        "url"  : "{0}",
        "name" : "Send Welcome Email",
        "parameters" : {{
            "event"    : "Site Creation",
            "product"  : "SharePoint Online"
        }}
        }}
    ]
    }}
    '

    if( -not ($siteScript = Get-PnPSiteScript | Where-Object -Property "Title" -eq "Send Site Creation Notificiation" ) )
    {
        Write-Host "Provisioning Site Script: Send Site Creation Notificiation"

        $schema = $template -f $powerAutomateOrLogicAppTriggerUrl

        $siteScript = Add-PnPSiteScript `
                                -Title       "Send Site Creation Notificiation"  `
                                -Description "Sends the site owners a welcome notice" `
                                -Content     $schema
    }


# create the site designs

if( -not (Get-PnPSiteDesign | Where-Object -Property "Title" -eq "Send Site Creation Notificiation" ) )
{
    Write-Host "Provisioning Site Design: Send Site Creation Notificiation"

    $design = Add-PnPSiteDesign `
                    -Title           "Send Site Creation Notificiation"  `
                    -Description     "Sends the site owners a welcome notice" `
                    -ThumbnailUrl    $previewImageUrl `
                    -SiteScriptIds   $siteScript.Id `
                    -WebTemplate     "TeamSite" `
                    -IsDefault

    Grant-PnPSiteDesignRights `
        -Identity $design.Id `
        -Principals "joe.rodgers@josrod.onmicrosoft.com", "c:0t.c|tenant|986b904f-0de9-416d-9fd9-7e5d8402e7c0" `
        -Rights View
}

<# 

# Remove Solution Commands

    Get-PnPSiteDesign | Where-Object -Property "Title" -eq "Send Site Creation Notificiation" | Remove-PnPSiteDesign -Force
    Get-PnPSiteScript | Where-Object -Property "Title" -eq "Send Site Creation Notificiation" | Remove-PnPSiteScript -Force

#>
