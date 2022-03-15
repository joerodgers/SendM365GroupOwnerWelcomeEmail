param
(
    $QueueItem, 
    $TriggerMetadata 
)

# setup telemetry

    $telemetry = New-Object Microsoft.ApplicationInsights.TelemetryClient
    $telemetry.InstrumentationKey = $env:APPINSIGHTS_INSTRUMENTATIONKEY
    $telemetry.TrackTrace( "Starting function app" )

# module import 

    Import-Module -Name "Microsoft.Graph.Authentication"
    Import-Module -Name "Microsoft.Graph.Groups"


# parameter config

    $clientId      = $env:SPO_CLIENTID
    $thumbprint    = $env:SPO_THUMBPRINT
    $tenantId      = $env:SPO_TENANTID
    $logicAppUrl   = $env:SEND_EMAIL_ENDPOINT_URI
    $failureEmail  = $env:FAILURE_EMAIL_ADDRESS
    $groupId       = $QueueItem.GroupId
    $siteUrl       = $QueueItem.SiteUrl

    $telemetry.TrackTrace( "Function configuration: ClientId: $clientId Thumbprint: $thumbprint TenantId: $tenantId LogicAppUrl: $logicAppUrl FailureEmail: $failureEmail" )
    $telemetry.TrackTrace( "Parameter configuration: GroupId: $groupId SiteUrl: $siteUrl" )

    $eventProperties = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    $eventProperties.Add( "GroupId", $groupId )
    $eventProperties.Add( "SiteUrl", $siteUrl )
    $eventProperties.Add( "OwnerCount", 0 )


    if( [string]::IsNullOrWhiteSpace($groupId) )
    {
        $telemetry.TrackTrace( "GroupId is empty, exiting." )
        $telemetry.TrackEvent( "NotificationIgnored", $eventProperties )
        return
    }

# connect to Microsoft Graph

    try
    {
        $null = Connect-MgGraph `
                    -ClientId              $clientId `
                    -CertificateThumbprint $thumbprint `
                    -TenantId              $tenantId `
                    -ErrorAction           Stop

        $telemetry.TrackTrace( "Connected to Microsoft Graph")
    }
    catch
    {
        Write-Error "Failed to connect to Microsoft Graph. Exception: $_"
        $telemetry.TrackException( $_.Exception )
    }


# get group owners

    try
    {
        $telemetry.TrackTrace( "Querying Microsoft Graph for GroupId $($groupId)" )

        $owners = @(Get-MgGroupOwner -GroupId $groupId -All)

        $eventProperties.OwnerCount = $owners.Count

        $telemetry.TrackTrace( "Retrieved $($owners.Count) group owners from Microsoft Graph" )
    }
    catch
    {
        $telemetry.TrackException( $_.Exception )
    }

# parse group owners

    if( $owners.Count -gt 0 )
    {
        $toAddresses = ($owners.AdditionalProperties).mail -join ";"
    }
    elseif( -not [string]::IsNullOrWhiteSpace($failureEmail) )
    {
        $toAddresses = $failureEmail
    }


# send email

    if( -not [string]::IsNullOrWhiteSpace($toAddresses) )
    {
        $telemetry.TrackTrace( "Sending email notification to: $toAddresses" )

        $json = [PSCustomObject] @{ 
                    OwnerEmailAddresses = $toAddresses
                    SiteUrl             = $siteUrl
                } | ConvertTo-Json -Depth 3

        try 
        {
            Invoke-RestMethod `
                -Method      "POST" `
                -Uri         $logicAppUrl `
                -ContentType "application/json" `
                -Body        $json `
                -ErrorAction Stop

            $telemetry.TrackTrace( "Notification detail: $json" )
            $telemetry.TrackEvent( "NotificationSent", $eventProperties )
        }
        catch
        {
            Write-Error "Failed to send email for group $($QueueItem.GroupId). Exception: $($_)"
            $telemetry.TrackEvent( "NotificationFailed", $eventProperties )
            $telemetry.TrackException( $_.Exception )
        }
    }
