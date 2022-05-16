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

    $clientId       = $env:SPO_CLIENTID
    $thumbprint     = $env:SPO_THUMBPRINT
    $tenantId       = $env:SPO_TENANTID
    $logicAppUrl    = $env:SEND_EMAIL_ENDPOINT_URI
    $failureEmail   = @($env:FAILURE_EMAIL_ADDRESS -split ";")
    $groupId        = $QueueItem.GroupId
    $siteUrl        = $QueueItem.SiteUrl
    $pilotEmails    = @($env:PILOT_EMAIL_ADDRESSES -split ";")
    $productionDate = [DateTime]::Parse( $env:PRODUCTION_DATE )

    $telemetry.TrackTrace( "Function configuration:
                                ClientId:       $clientId
                                Thumbprint:     $thumbprint
                                TenantId:       $tenantId
                                LogicAppUrl:    $logicAppUrl
                                FailureEmail:   $failureEmail
                                PilotEmails:    $($pilotEmails -join ';')
                                ProductionDate: $productionDate" )
 
    $telemetry.TrackTrace( "Execution Parameters:
                                GroupId: $groupId
                                SiteUrl: $siteUrl" )


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

    $emails = @()

    if( [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Exists )
    {
        try
        {
            $group  = Get-MgGroup -GroupId $groupId 
            $emails = @(Get-MgGroupOwner -GroupId $groupId -All -Property mail).AdditionalProperties.mail
   
            $telemetry.TrackTrace( "Retrieved $($emails.Count) group owners from Microsoft Graph" )
            $eventProperties.OwnerCount = $emails.Count
        }
        catch
        {
            $telemetry.TrackException( $_.Exception )
        }        
    }


# short circut during pre-produciton phase

    if( [DateTime]::Now -lt $productionDate -and $pilotEmails.Count -gt 0 -and $emails.Count -gt 0 )
    {
        $temp = @()

        foreach( $email in $emails )
        {
            if( $pilotEmails -contains $email.Trim() )
            {
                $temp += $email.Trim()
            }
            else 
            {
                $telemetry.TrackTrace( "Removing non-preview group owner email: $email" )
            }
        }

        $emails = $temp
    }


# parse group owners

    if( $emails.Count -eq 0 -and $failureEmail.Count -gt 0 )
    {
        $telemetry.TrackTrace( "Failing back to default email address for group: $groupId" )
        $emails = $failureEmail
    }


# send email

    if( $emails.Count -gt 0 )
    {
        $json = [PSCustomObject] @{ 
                    OwnerEmailAddresses = ($emails -join ";")
                    SiteUrl             = $siteUrl
                    DisplayName         = $group.DisplayName
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
