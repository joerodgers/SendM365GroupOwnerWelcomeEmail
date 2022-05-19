#requires -modules "PnP.PowerShell"

param
(
    [Parameter(Mandatory=$false)]
    [string]
    $Tenant,

    [Parameter(Mandatory=$true)]
    [Guid]
    $ClientId,

    [Parameter(Mandatory=$true)]
    [string]
    $CertificateThumbprint,

    [Parameter(Mandatory=$true)]
    [string]
    $WebHookUrl,

    [Parameter(Mandatory=$true)]
    [string[]]
    $Principal
)

[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol   = [System.Net.SecurityProtocolType]::Tls12

# connect to tenant admin

    Connect-PnPOnline `
        -Url        "https://$Tenant-admin.sharepoint.com" `
        -ClientId   $ClientId `
        -Thumbprint $CertificateThumbprint `
        -Tenant     "$Tenant.onmicrosoft.com"


# create site script 

    $title    = "Send Site Creation Notificiation"
    $desc     = "Sends the site owners a welcome notice"
    $template = '
    {{
    "$schema" : "schema.json",
    "actions" : [
        {{
        "verb" : "triggerFlow",
        "url"  : "{0}",
        "name" : "Trigger Webhook",
        "parameters" : {{
            "event"    : "Site Creation",
            "product"  : "SharePoint Online"
        }}
        }}
    ]
    }}
    '

    if( $siteScript = Get-PnPSiteScript | Where-Object -Property "Title" -eq $title )
    {
        Write-Host "Removing existing site script: '$($siteDesign.Title)'"
        Remove-PnPSiteScript -Identity $siteScript.Id -Force
    }    

    Write-Host "Provisioning Site Script: $title"

    $siteScript = Add-PnPSiteScript `
                        -Title       $title `
                        -Description $desc `
                        -Content     ($template -f $WebHookUrl)




# create the site template

    if( $siteDesign = Get-PnPSiteDesign | Where-Object -Property "Title" -eq $title )
    {
        Write-Host "Removing existing site design: '$($siteDesign.Title)'"
        Remove-PnPSiteDesign -Identity $siteDesign.Id -Force
    }    

    Write-Host "Provisioning Site Design: $title"

    $design = Add-PnPSiteDesign `
                    -Title           $title  `
                    -Description     $desc `
                    -SiteScriptIds   $siteScript.Id `
                    -WebTemplate     "TeamSite" `
                    -IsDefault

# set acl

    if( $PSBoundParameters.ContainsKey( "Principal" ) )
    {
        Grant-PnPSiteDesignRights `
                    -Identity   $design.Id `
                    -Principals $Principal `
                    -Rights     View
    }

<# 

# remove script and template commands

    Get-PnPSiteDesign | Where-Object -Property "Title" -eq "Send Site Creation Notificiation" | Remove-PnPSiteDesign -Force
    Get-PnPSiteScript | Where-Object -Property "Title" -eq "Send Site Creation Notificiation" | Remove-PnPSiteScript -Force

#>
