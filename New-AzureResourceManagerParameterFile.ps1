param
(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Production", "Test", "Development")]
    [string]
    $Environment,

    [Parameter(Mandatory=$false)]
    [switch]
    $Force
)

Import-Module -Name "$PSScriptRoot\resources\resources.psm1" -Force -ErrorAction Stop

New-ParameterFile `
        -Environment $Environment `
        -Force:$Force.IsPresent
