param
(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Production", "Test", "Development")]
    [string]
    $Environment
)

Import-Module -Name "$PSScriptRoot\resources\resources.psm1" -Force -ErrorAction Stop

$parameters_production = @{
    "clientId"               = $env:O365_CLIENTID
    "tenantId"               = $env:O365_TENANTID
    "certificateThumbprint"  = $env:O365_THUMBPRINT
    "certificatePfxPassword" = $env:O365_CERT_PWD
    "certificatePfxBase64"   = [System.Convert]::ToBase64String(( Get-Content -Path "$PSScriptRoot\resources\certificate.pfx" -Raw -Encoding Byte ))
    "mailboxAddress"         = ""
    "emailSubject"           = ""
    "fallbackEmailAddress"   = ""
    "productionDate"         = ""
    "pilotEmailAddresses"    = ""
}


$parameters_test = @{
    "clientId"               = $env:O365_CLIENTID
    "tenantId"               = $env:O365_TENANTID
    "certificateThumbprint"  = $env:O365_THUMBPRINT
    "certificatePfxPassword" = $env:O365_CERT_PWD
    "certificatePfxBase64"   = [System.Convert]::ToBase64String(( Get-Content -Path "$PSScriptRoot\resources\certificate.pfx" -Raw -Encoding Byte ))
    "mailboxAddress"         = ""
    "emailSubject"           = ""
    "fallbackEmailAddress"   = ""
    "productionDate"         = ""
    "pilotEmailAddresses"    = ""
}

$parameters_development = @{
    "clientId"               = $env:O365_CLIENTID
    "tenantId"               = $env:O365_TENANTID
    "certificateThumbprint"  = $env:O365_THUMBPRINT
    "certificatePfxPassword" = $env:O365_CERT_PWD
    "certificatePfxBase64"   = [System.Convert]::ToBase64String(( Get-Content -Path "$PSScriptRoot\resources\certificate.pfx" -Raw -Encoding Byte ))
    "mailboxAddress"         = ""
    "emailSubject"           = ""
    "fallbackEmailAddress"   = ""
    "productionDate"         = ""
    "pilotEmailAddresses"    = ""
}

switch( $Environment )
{
    "Production"
    {
        $parameters = $parameters_production 
    }
    "Test"
    {
        $parameters = $parameters_test 
    }
    "Development"
    {
        $parameters = $parameters_development 
    }
}

Set-DefaultParameterFileValue `
    -Path  "$PSScriptRoot\resources\azure-deploy-parameters.$Environment.json" `
    -Value $parameters

