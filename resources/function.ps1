param
(
    $QueueItem, 
    $TriggerMetadata 
)

if( $QueueItem -is [string] )
{
    $QueueItem = $QueueItem | ConvertFrom-Json -Depth 100
}

Import-Module -Name "Microsoft.Graph.Authentication"
Import-Module -Name "Microsoft.Graph.Groups"

<#

    $QueueItem - Parameter is a HashTable sent from the the JSON message dropped in the Azure Storage Queue

        $QueueItem.GroupId = "10d1f773-ed11-49e1-bf8c-e32343db39d8"
        $QueueItem.SiteUrl = "https://tenant.sharepoint.com/sites/sitename"

    $TriggerMetadata - Parameter is used to supply additional information about the trigger. See https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference-powershell?tabs=portal#triggermetadata-parameter
    
#>

# credentials

    $clientId      = $env:SPO_CLIENTID
    $thumbprint    = $env:SPO_THUMBPRINT
    $tenantId      = $env:SPO_TENANTID
    $logicAppUrl   = $env:SEND_EMAIL_ENDPOINT_URI
    $failureEmail  = $env:FAILURE_EMAIL_ADDRESS
    $groupId       = $QueueItem.GroupId

# connect to Microsoft Graph

    Connect-MgGraph `
        -ClientId              $clientId `
        -CertificateThumbprint $thumbprint `
        -TenantId              $tenantId | Out-Null

# get group owners

    $owners = (Get-MgGroupOwner -GroupId $groupId -All  Select-Object -ExpandProperty "AdditionalProperties").mail -join ";"

    if( [string]::IsNullOrWhiteSpace($owners) )
    {
        if( -not [string]::IsNullOrWhiteSpace($failureEmail) )
        {
            Write-Warning "Group owners for group '$groupId' was found using the fallback address"
            $owners = $failureEmail
        }
        else
        {
            Write-Error "Group owners for group '$groupId' was found and no fallback email was defined."
            return
        }
    }

# send email

    $json = [PSCustomObject] @{ 
                OwnerEmailAddresses = $owners
                SiteUrl             = $QueueItem.SiteUrl
            } | ConvertTo-Json -Depth 100

    try 
    {
        Invoke-RestMethod `
            -Method      "POST" `
            -Uri         $logicAppUrl `
            -ContentType "application/json" `
            -Body        $json `
            -ErrorAction Stop
    }
    catch
    {
        Write-Error "Failed to send email for group $($QueueItem.GroupId). Exception: $($_)"
    }
