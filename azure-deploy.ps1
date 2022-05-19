#requires -modules "Az.Resources", "Az.Accounts", "Az.Websites"

param
(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Production", "Test", "Development")]
    [string]
    $Environment = "Development",

    [Parameter(Mandatory=$true)]
    [Guid]
    $TenantId,

    [Parameter(Mandatory=$true)]
    [Guid]
    $SubscriptionId,

    [Parameter(Mandatory=$true)]
    [string]
    $ResourceGroup,

    [Parameter(Mandatory=$false)]
    [string]
    $Location = "eastus"
)

[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol   = [System.Net.SecurityProtocolType]::Tls12   

$ctx = Get-AzContext

if( $ctx.Tenant.Id -ne $TenantId.ToString() -or $ctx.Subscription.SubscriptionId -ne $SubscriptionId.ToString() )
{
    Write-Host "[$(Get-Date)] - Prompting for Azure credentials"
    Login-AzAccount -Tenant $TenantId -WarningAction SilentlyContinue

    $ctx = Get-AzContext
}

$subscription = Select-AzSubscription -Subscription $SubscriptionId -WarningAction SilentlyContinue

Write-Host "[$(Get-Date)] - Connected as: $($ctx.Account.Id)"
Write-Host "[$(Get-Date)] - Subscription: $($subscription.Subscription.Name)"
Write-Host "[$(Get-Date)] - Environment:  $Environment"

$templatePath      = Join-Path -Path $PSScriptRoot -ChildPath "resources\azure-deploy.json"
$parameterPath     = Join-Path -Path $PSScriptRoot -ChildPath "resources\azure-deploy-parameters.$Environment.json"
$requirementsPath  = Join-Path -Path $PSScriptRoot -ChildPath "resources\requirements.psd1"
$functionPath      = Join-Path -Path $PSScriptRoot -ChildPath "resources\function.ps1"

# path validation

    foreach( $path in @($templatePath, $parameterPath, $requirementsPath, $functionPath) )
    {
        if( -not (Test-Path -Path $path -PathType Leaf) )
        {
            Write-Error "Required file not found: $path"
            return
        }
    }

    if( -not (Test-Path -Path "$PSScriptRoot\deploymentlogs" -PathType Container) )
    {
        New-Item -ItemType Directory -Path "$PSScriptRoot\deploymentlogs"
    }

    Write-Host "[$(Get-Date)] - Template Path:     $templatePath"
    Write-Host "[$(Get-Date)] - Parameter Path:    $parameterPath"
    Write-Host "[$(Get-Date)] - Requirements Path: $requirementsPath"
    Write-Host "[$(Get-Date)] - Function Path:     $functionPath"
    Write-Host "[$(Get-Date)] - Log Path:          $PSScriptRoot\deploymentlogs"


    if( -not (Get-AzResourceGroup -Name $ResourceGroup -Location $Location -ErrorAction SilentlyContinue) )
    {
        Write-Error "Resource group not found: $ResourceGroup"
        return
    }

# start deployment

    Write-Host 
    Write-Host "[$(Get-Date)] - Starting deployment"
     
    $deployment = New-AzResourceGroupDeployment `
                        -ResourceGroupName     $ResourceGroup `
                        -TemplateFile          $templatePath `
                        -TemplateParameterFile $parameterPath

    Write-Host "[$(Get-Date)] - Deployment $($deployment.ProvisioningState)"

    $deployment.OutputsString | Set-Content -Path "$PSScriptRoot\deploymentlogs\deploymentoutput_$(Get-Date -Format FileDateTime).log"

    if( $deployment.ProvisioningState -ne "Succeeded" ) { return }



# upload requirements.psd1 to configure modules for the Azure function

    Write-Host "[$(Get-Date)] - Deploying requirements.psd1"

    Compress-Archive `
                -Path            $requirementsPath `
                -DestinationPath "$PSScriptRoot\resources\requirements.zip" `
                -Force

    $null = Publish-AzWebApp `
                -ResourceGroupName $resourceGroup `
                -Name              $deployment.Outputs.functionAppName.Value `
                -ArchivePath       "$PSScriptRoot\resources\requirements.zip" `
                -Force



# upload function.ps1 for the Azure function

    Write-Host "[$(Get-Date)] - Deploying function.ps1"

    Compress-Archive `
                -Path            $functionPath `
                -DestinationPath "$PSScriptRoot\resources\function.zip" `
                -Force

    $null = Publish-AzWebApp `
                -ResourceGroupName $resourceGroup `
                -Name              $deployment.Outputs.functionAppName.Value `
                -ArchivePath       "$PSScriptRoot\resources\function.zip" `
                -Force



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

