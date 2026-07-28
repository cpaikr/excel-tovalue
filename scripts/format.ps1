[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$solutionPath = Join-Path $repositoryRoot 'ExcelFormulaRemover.slnx'

function Invoke-NativeCommand {
    param(
        [Parameter(Mandatory)]
        [string] $FilePath,

        [string[]] $ArgumentList = @()
    )

    Write-Host "> $FilePath $($ArgumentList -join ' ')"
    & $FilePath @ArgumentList
    $exitCode = $LASTEXITCODE

    if ($exitCode -ne 0) {
        throw "Command '$FilePath' failed with exit code $exitCode."
    }
}

Push-Location $repositoryRoot
try {
    Invoke-NativeCommand -FilePath 'dotnet' -ArgumentList @(
        'restore',
        $solutionPath,
        '--locked-mode'
    )

    Invoke-NativeCommand -FilePath 'dotnet' -ArgumentList @(
        'format',
        $solutionPath,
        '--no-restore',
        '--verbosity',
        'minimal'
    )

    Invoke-NativeCommand -FilePath 'rumdl' -ArgumentList @(
        'fmt',
        '--respect-gitignore=true',
        $repositoryRoot
    )

    Write-Host 'Repository formatting completed successfully.'
}
finally {
    Pop-Location
}
