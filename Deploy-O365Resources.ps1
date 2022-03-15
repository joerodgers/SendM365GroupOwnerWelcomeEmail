#requires -modules "PnP.PowerShell"

[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol   = [System.Net.SecurityProtocolType]::Tls12

$clientId   = $env:O365_CLIENTID
$thumbprint = $env:O365_THUMBPRINT
$tenantId   = $env:O365_TENANTID
$tenant     = $env:O365_TENANT

# you will see this as output in the Deploy-AzureResources.ps1 script or pull directly from the Logic App's trigger action.
$powerAutomateOrLogicAppTriggerUrl = "https://prod-50.eastus.logic.azure.com:443/xxxxx"

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
        "name" : "Apply Site Template",
        "parameters" : {{
            "event"    : "Site Creation",
            "product"  : "SharePoint Online"
        }}
        }}
    ]
    }}
    '

    if( -not ($siteScript = Get-PnPSiteScript | Where-Object -Property "Title" -eq "Public Site Approval Workflow") )
    {
        Write-Host "Provisioning Site Script: Public Site Approval Workflow"

        $schema = $template -f $powerAutomateOrLogicAppTriggerUrl

        $siteScript = Add-PnPSiteScript `
                                -Title       "Public Site Approval Workflow" `
                                -Description "Initiates an approval workflow to convert a site from private to public." `
                                -Content     $schema
    }


# create the site designs

if( -not (Get-PnPSiteDesign | Where-Object -Property "Title" -eq "Public Site Approval Workflow" ) )
{
    Write-Host "Provisioning Site Design: Public Site Approval Workflow"

    Add-PnPSiteDesign `
        -Title           "Public Site Approval Workflow" `
        -Description     "The template initiates an approval workflow for converting a SharePoint Online site from private to public." `
        -ThumbnailUrl    $previewImageUrl `
        -SiteScriptIds   $siteScript.Id `
        -WebTemplate     "TeamSite"
}


<# 

# Remove Solution Commands

    Get-PnPSiteDesign | Where-Object -Property "Title" -eq "Public Site Approval Workflow" | Remove-PnPSiteDesign -Force
    Get-PnPSiteScript | Where-Object -Property "Title" -eq "Public Site Approval Workflow" | Remove-PnPSiteScript -Force

#>
