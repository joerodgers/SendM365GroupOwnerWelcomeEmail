param
(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Production", "Test", "Development")]
    [string]
    $Environment
)

Import-Module -Name "$PSScriptRoot\resources\resources.psm1" -Force -ErrorAction Stop

if( $PSVersionTable.PSVersion.Major -le 5 )
{
    $splat = @{ Raw = $true; Encoding = "Byte" }
}
else
{
    $splat = @{ Raw = $true; AsByteStream = $true }
}

$parameters_development = @{
    "clientId"               = $env:O365_CLIENTID
    "tenantId"               = $env:O365_TENANTID
    "certificateThumbprint"  = $env:O365_THUMBPRINT
    "certificatePfxPassword" = $env:O365_CERT_PWD
    "certificatePfxBase64"   = [System.Convert]::ToBase64String(( Get-Content -Path "$PSScriptRoot\resources\certificate.pfx" @splat ))
    "mailboxAddress"         = "support@contoso.com"
    "emailSubject"           = "Information about your new group"
    "emailBody"              = (Get-Content -Path "$PSScriptRoot\resources\email_template.html" -Raw)
    "fallbackEmailAddress"   = "support@contoso.com"
    "productionDate"         = "12/21/2025 18:00:00"
    "pilotEmailAddresses"    = "john.doe@contoso.com;jane.doe@contoso.com"
}

$parameters_test = @{
    "clientId"               = ""
    "tenantId"               = ""
    "certificateThumbprint"  = ""
    "certificatePfxPassword" = ""
    "certificatePfxBase64"   = "" # [System.Convert]::ToBase64String(( Get-Content -Path "$PSScriptRoot\resources\certificate.pfx" @splat ))
    "mailboxAddress"         = "support@contoso.com"
    "emailSubject"           = "Information about your new group"
    "emailBody"              = (Get-Content -Path "$PSScriptRoot\resources\email_template.html" -Raw)
    "fallbackEmailAddress"   = "support@contoso.com"
    "productionDate"         = "12/21/2025 18:00:00"
    "pilotEmailAddresses"    = "john.doe@contoso.com;jane.doe@contoso.com"
}

$parameters_production = @{
    "clientId"               = ""
    "tenantId"               = ""
    "certificateThumbprint"  = ""
    "certificatePfxPassword" = ""
    "certificatePfxBase64"   = "" # [System.Convert]::ToBase64String(( Get-Content -Path "$PSScriptRoot\resources\certificate.pfx" @splat ))
    "mailboxAddress"         = "support@contoso.com"
    "emailSubject"           = "Information about your new group"
    "emailBody"              = (Get-Content -Path "$PSScriptRoot\resources\email_template.html" -Raw)
    "fallbackEmailAddress"   = "support@contoso.com"
    "productionDate"         = "12/21/2025 18:00:00"
    "pilotEmailAddresses"    = "john.doe@contoso.com;jane.doe@contoso.com"
}

$parameters = switch( $Environment )
              {
                  "Development" { $parameters_development }
                  "Test"        { $parameters_test        }
                  "Production"  { $parameters_production  }
              }

Set-DefaultParameterFileValue `
    -Path  "$PSScriptRoot\resources\azure-deploy-parameters.$Environment.json" `
    -Value $parameters

