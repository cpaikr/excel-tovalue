[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$solutionPath = Join-Path $repositoryRoot 'ExcelFormulaRemover.slnx'
$testResultsDirectory = Join-Path $repositoryRoot 'artifacts/test-results'

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

function Test-HasVulnerabilities {
    param(
        [AllowNull()]
        [object] $Value
    )

    if ($null -eq $Value -or $Value -is [string] -or $Value -is [ValueType]) {
        return $false
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [pscustomobject]) {
        foreach ($item in $Value) {
            if (Test-HasVulnerabilities -Value $item) {
                return $true
            }
        }

        return $false
    }

    foreach ($property in $Value.PSObject.Properties) {
        if ($property.Name -eq 'vulnerabilities' -and @($property.Value).Count -gt 0) {
            return $true
        }

        if (Test-HasVulnerabilities -Value $property.Value) {
            return $true
        }
    }

    return $false
}

function Invoke-VulnerabilityAudit {
    $arguments = @(
        'list',
        $solutionPath,
        'package',
        '--vulnerable',
        '--include-transitive',
        '--format',
        'json',
        '--no-restore'
    )

    Write-Host "> dotnet $($arguments -join ' ')"
    $auditLines = & dotnet @arguments
    $exitCode = $LASTEXITCODE
    $auditText = $auditLines -join [Environment]::NewLine

    if ($auditText.Length -gt 0) {
        Write-Host $auditText
    }

    if ($exitCode -ne 0) {
        throw "Package vulnerability audit failed with exit code $exitCode."
    }

    if ([string]::IsNullOrWhiteSpace($auditText)) {
        throw 'Package vulnerability audit returned no JSON output.'
    }

    try {
        $auditResult = $auditText | ConvertFrom-Json
    }
    catch {
        throw "Package vulnerability audit did not return valid JSON: $($_.Exception.Message)"
    }

    if (Test-HasVulnerabilities -Value $auditResult) {
        throw 'The package vulnerability audit found one or more vulnerable dependencies.'
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
        '--verify-no-changes',
        '--no-restore',
        '--verbosity',
        'minimal'
    )

    Invoke-NativeCommand -FilePath 'dotnet' -ArgumentList @(
        'build',
        $solutionPath,
        '--configuration',
        'Release',
        '--no-restore',
        '-p:ContinuousIntegrationBuild=true'
    )

    New-Item -ItemType Directory -Path $testResultsDirectory -Force | Out-Null

    Invoke-NativeCommand -FilePath 'dotnet' -ArgumentList @(
        'test',
        $solutionPath,
        '--configuration',
        'Release',
        '--no-build',
        '--no-restore',
        '--logger',
        'console;verbosity=normal',
        '--logger',
        'trx',
        '--results-directory',
        $testResultsDirectory
    )

    Invoke-VulnerabilityAudit

    Invoke-NativeCommand -FilePath 'rumdl' -ArgumentList @(
        'check',
        '--respect-gitignore=true',
        $repositoryRoot
    )

    Write-Host 'Repository checks completed successfully.'
}
finally {
    Pop-Location
}
