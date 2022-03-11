function Format-Json
{
    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [String]
        $json,

        [Parameter(Mandatory=$false)]
        [int]
        $NumberSpaces = 4
    ) 

    begin
    {
        $sb = New-Object System.Text.StringBuilder

        $indent = 0;
    }
    end
    {
        $lines =  $json -Split [System.Environment]::NewLine

        foreach( $line in $lines )
        {
            if( $line -match '[\}\]]' ) 
            {
                $indent--
            }
      
            $null = $sb.AppendFormat( "{0}{1}{2}", (' ' * $indent * $NumberSpaces), $line.TrimStart().Replace(':  ', ': '), [System.Environment]::NewLine )
             
            if ($line -match '[\{\[]')
            {
                $indent++
            }
        }

        $sb.ToString()
    }
}

function New-ParameterFile 
{
    [CmdletBinding()]
    param 
    (
        [parameter(Mandatory=$false)]
        [string] 
        $InputPath = (Join-Path $PSScriptRoot "azure-deploy.json"),
        
        [parameter(Mandatory=$false)]
        [string] 
        $OutputPath = (Join-Path $PSScriptRoot "azure-deploy-paramters.json"),

        [parameter(Mandatory=$false)]
        [switch] 
        $SkipOptionalParameters,

        [parameter(Mandatory=$false)]
        [ValidateSet("Production", "Test", "Development")]
        [string[]] 
        $Environment,

        [parameter(Mandatory=$false)]
        [switch] 
        $Force
    )
    begin
    {
        # shell object
        $parameterObject =  [PSCustomObject] @{
                                '$schema'      = "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#"
                                contentVersion = "1.0.0.0"
                                parameters     = $null
                            }
    
        $defaultValue = [PSCustomObject] @{ value = "Prompt" }
        $parameters   = [PSCustomObject] @{ parameters = @() }
    }
    process
    {

        # read in the template and convert to an object
        $template = Get-Content -Path $InputPath -Raw -ErrorAction Stop | ConvertFrom-Json
        
        $param = [PSCustomObject] @{}

        # get the parameters section
        foreach( $parameter in $template.parameters.psobject.members | Where-Object -Property "MemberType" -eq "NoteProperty" )
        {
            $isRequired = ($parameter.value | Get-Member | Where-Object -Property "Name" -eq "defaultValue") -eq $null

            if( $SkipOptionalParameters.IsPresent -and -not $isRequired )
            {
                continue
            }

            $param | Add-Member -MemberType NoteProperty -Name $parameter.Name -Value $defaultValue
        }

        $parameterObject.parameters = $param

        if( $PSBoundParameters.ContainsKey( "Environment" ) )
        {
            foreach( $env in $Environment )
            {
                $environmentOutputPath = $OutputPath -replace '.json', ".$env.json".ToLower()

                if( (Test-Path -Path $environmentOutputPath -PathType Leaf) -and -not $Force.IsPresent )
                {
                    Write-Error "Existing template found at $($OutputPath).  Use -Force to overwrite."
                    continue
                }

                $parameterObject | ConvertTo-Json -Depth 100 | Format-Json | Set-Content -Path $environmentOutputPath
            }

            return
        }

        if( (Test-Path -Path $OutputPath -PathType Leaf) -and -not $Force.IsPresent )
        {
            Write-Error "Existing template found at $($OutputPath).  Use -Force to overwrite."
            return
        }

        $parameterObject | ConvertTo-Json -Depth 100 | Format-Json | Set-Content -Path $OutputPath
    }
    end
    {
    }
}

New-ParameterFile `
    -InputPath  "$PSScriptRoot\resources\azure-deploy.json" `
    -OutputPath "$PSScriptRoot\resources\azure-deploy-parameters.json" `
    -Environment Production, Development, Test
