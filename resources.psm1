function Show-OAuthWindow
{
    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory=$true)]
        [System.Uri]
        $Url
    )

    begin
    {
        Add-Type -AssemblyName System.Windows.Forms
    }
    process 
    {
        $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{
            Width  = 420
            Height = 600
            Url    = $Url
        }
    
        $web.ScriptErrorsSuppressed = $true
    
        $web.Add_DocumentCompleted( {
                if ($web.Url.AbsoluteUri -match "error=[^&]*|code=[^&]*") { $form.Close() }
            })

        $form = New-Object -TypeName System.Windows.Forms.Form -Property @{
            Width  = 440
            Height = 640
        }
    
        $form.Controls.Add($web)
    
        $form.Add_Shown( {
                $form.BringToFront()
                $null = $form.Focus()
                $form.Activate()
                $web.Navigate($Url)
            })

        $null = $form.ShowDialog()

        $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
        
        $output = @{}
        
        foreach ($key in $queryOutput.Keys) 
        {
            $output["$key"] = $queryOutput[$key]
        }

        [pscustomobject]$output
    }
}