#requires -modules "Az.Resources", "Az.Accounts", "Az.Websites"

[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol   = [System.Net.SecurityProtocolType]::Tls12   

Import-Module -Name "$PSScriptRoot\resources.psm1" -ErrorAction Stop -Force

$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop


#Login-AzAccount -Tenant $env:MSFT_TENANTID -WarningAction SilentlyContinue
#Select-AzSubscription -Subscription $env:MSFT_SUBSCRIPTIONID -WarningAction SilentlyContinue

$resourceGroup     = "RG-SPOWELCOMEEMAIL-NPROD2-USEAST"
$templatePath      = Join-Path -Path $PSScriptRoot -ChildPath "resources\azure-deploy.json"
$parameterPath     = Join-Path -Path $PSScriptRoot -ChildPath "resources\azure-deploy-parameters.development.json"
$requirementsPath  = Join-Path -Path $PSScriptRoot -ChildPath "resources\requirements.psd1"

# start deployment

    Write-Host "[$(Get-Date)] - Starting deployment"
     
    $deployment = New-AzResourceGroupDeployment `
        -ResourceGroupName     $resourceGroup `
        -TemplateFile          $templatePath `
        -TemplateParameterFile $parameterPath

    Write-Host "[$(Get-Date)] - Deployment $($deployment.ProvisioningState)"

    $deployment.OutputsString | Set-Content -Path "$PSScriptRoot\deploymentlogs\deploymentoutput_$(Get-Date -Format FileDateTime).log"

    if( $deployment.ProvisioningState -ne "Succeeded" ) { return }


# upload requirements.psd1 to configure modules for the Azure function

    Write-Host "[$(Get-Date)] - Deploying requirements.psd1"

    Compress-Archive `
                -Path $requirementsPath `
                -DestinationPath "$PSScriptRoot\resources\requirements.zip" `
                -Force

    $null = Get-AzWebApp `
                -ResourceGroupName $resourceGroup `
                -Name              $deployment.Outputs.functionAppName.Value
    
    $null = Publish-AzWebApp `
                -ResourceGroupName $resourceGroup `
                -Name              $deployment.Outputs.functionAppName.Value `
                -ArchivePath       "$PSScriptRoot\resources\requirements.zip" `
                -Force

#>

# authorize api connections


    Write-Host "[$(Get-Date)] - Authorizing API Connections"

    $connections = Get-AzResource -ResourceType "Microsoft.Web/connections" -ResourceGroupName $resourceGroup 

    $authorizedConnections = @()

    foreach( $connection in $connections )
    {
        $conn = Get-AzResource -ResourceId $connection.ResourceId
        
        $connectionStatus = [PSCustomObject] @{
                                Connection = $conn.Name
                                Status     = $conn.Properties.statuses[0].status
                            }

        if( $conn.Properties.statuses[0].status -ne "connected" )
        {
            Write-Host "[$(Get-Date)] - Authorizing Connection: $($conn.Name)"

            $parameters = @{
                "parameters" = ,@{
                    "parameterName" = "token";
                    "redirectUrl"   = "https://ema1.exp.azure.com/ema/default/authredirect"
                }
            }

            $consentLink = Invoke-AzResourceAction `
                                    -Action    "listConsentLinks" `
                                    -ResourceId $conn.ResourceId `
                                    -Parameters $parameters `
                                    -Force    
        
            $consentCode = Show-OAuthWindow -Url $consentLink.value.link
        
            try
            {
                Invoke-AzResourceAction `
                        -Action      "confirmConsentCode" `
                        -ResourceId  $conn.ResourceId `
                        -Parameters  @{ "code" = $consentCode.code } `
                        -Force `
                        -ErrorAction Stop
            }
            catch
            {
                # throws the following error due to lack of response:
                # Cannot process argument because the value of argument "obj" is null. Change the value of argument "obj" to a non-null value

                if( $_ -notmatch 'Cannot process argument because the value of argument "obj" is null' )
                {
                    throw $_
                }
            }
            
            $conn = Get-AzResource -ResourceId $connection.ResourceId

            $connectionStatus.Status = $conn.Properties.statuses[0].status
        }

        $authorizedConnections += $connectionStatus
    }

    $connections = Get-AzResource -ResourceType "Microsoft.Web/connections" -ResourceGroupName $resourceGroup 

    foreach( $connection in $connections )
    {
        $conn = Get-AzResource -ResourceId $connection.ResourceId

        Write-Host "[$(Get-Date)] - `tConnection $($conn.Name): $($conn.Properties.statuses[0].status)"
    }

