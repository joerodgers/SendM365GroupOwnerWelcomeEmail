#requires -modules "Az.Resources", "Az.Accounts", "Az.Websites"

[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12   

# Login-AzAccount -Tenant "72f988bf-86f1-41af-91ab-2d7cd011db47"
# Select-AzSubscription -Subscription "d432671f-fd2d-449f-afdf-010ba093eace" -WarningAction SilentlyContinue

$resourceGroup     = "RG-SPOWELCOMEEMAIL-NPROD-USEAST"
$certPath          = Join-Path -Path $PSScriptRoot -ChildPath "resources\certificate.pfx"
$templatePath      = Join-Path -Path $PSScriptRoot -ChildPath "resources\azure-deploy.json"
$parameterPath     = Join-Path -Path $PSScriptRoot -ChildPath "resources\azure-deploy-parameters.development.json"
$functionPath      = Join-Path -Path $PSScriptRoot -ChildPath "resources\function.ps1"
$requirementsPath  = Join-Path -Path $PSScriptRoot -ChildPath "resources\requirements.psd1"
$emailTemplatePath = Join-Path -Path $PSScriptRoot -ChildPath "resources\email_template.html"


<#

$parameters = Get-Content -Path $parameterPath | ConvertFrom-Json
$parameters.parameters.clientId.value               = $env:O365_CLIENTID
$parameters.parameters.tenantId.value               = $env:O365_TENANTID
$parameters.parameters.certificateThumbprint.value  = $env:O365_THUMBPRINT
$parameters.parameters.certificatePfxPassword.value = 'pass@word1'
$parameters.parameters.certificatePfxBase64.value   = [System.Convert]::ToBase64String(( Get-Content -Path $certPath -Raw -Encoding Byte ))
$parameters.parameters.mailboxAddress.value         = 'josrod@microsoft.com'
$parameters.parameters.emailSubject.value           = 'Welcome to SharePoint Online'
$parameters.parameters.emailBody.value              = (Get-Content -Path $emailTemplatePath -Raw).ToString()
$parameters.parameters.fallbackEmailAddress.value   = 'josrod@microsoft.com'
$parameters.parameters.functionCode.value           = (Get-Content -Path $functionPath -Raw).ToString()
$parameters | ConvertTo-Json -Depth 100 | Set-Content -Path $parameterPath

#> 


$deploymentParameters = @{
    ResourceGroupName       = $resourceGroup
    TemplateFile            = $templatePath
    TemplateParameterFile   = $parameterPath
}

if( Test-Path -Path $templatePath -PathType Leaf )
{
    Write-Host "$(Get-Date) - Deploying Solution"

    $deployment = New-AzResourceGroupDeployment @deploymentParameters

    $deploymentResults = [PSCustomObject] @{
                             DeploymentName    = $deployment.DeploymentName
                             ResourceGroupName = $deployment.ResourceGroupName
                             ProvisioningState = $deployment.ProvisioningState
                             Timestamp         = $deployment.Timestamp.ToLocalTime().ToString("yyyy-MM-ddTHH-mm-ss")
                             Mode              = $deployment.Mode
                         }
    
    $deployment.Outputs.GetEnumerator() | ForEach-Object { $deploymentResults | Add-Member -MemberType NoteProperty -Name $_.Key -Value $_.value.value }

    New-Item -Path "$PSScriptRoot\deploymentlogs" -ItemType Directory -ErrorAction Ignore | Out-Null

    $deploymentResults | Export-Csv -Path "$PSScriptRoot\deploymentlogs\deploymentoutput_$($deploymentResults.Timestamp).log" -NoTypeInformation
}
#>

if( $deployment -and (Test-Path -Path $requirementsPath -PathType Leaf) )
{
    if( Get-AzWebApp -ResourceGroupName $resourceGroup -Name $deployment.Outputs.functionAppName.Value )
    {
        Write-Host "$(Get-Date) - Updating function requirements.psd1"
    
        Compress-Archive -Path $requirementsPath -DestinationPath "$PSScriptRoot\resources\requirements.zip" -Force

        $null = Publish-AzWebApp `
                    -ResourceGroupName $resourceGroup `
                    -Name              $deployment.Outputs.functionAppName.Value `
                    -ArchivePath       "$PSScriptRoot\resources\requirements.zip" `
                    -Force
    }
}

