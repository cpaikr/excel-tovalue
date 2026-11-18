# Certify a published CLI archive for the current platform: install it through
# the release installer, then run the process contract tests against the
# installed executable.
[CmdletBinding()]
param(
    # A release directory produced by publish.ps1.
    [Parameter(Mandatory, Position = 0)]
    [string] $ReleaseDirectory,

    [Parameter(Mandatory, Position = 1)]
    [string] $Version
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$testProject = Join-Path $repositoryRoot 'tests/ExcelFormulaRemover.Cli.Tests/ExcelFormulaRemover.Cli.Tests.csproj'
$ReleaseDirectory = (Resolve-Path -LiteralPath $ReleaseDirectory).ProviderPath
$checksums = Join-Path $ReleaseDirectory 'SHA256SUMS.txt'

if ($IsWindows) {
    $archive = Join-Path $ReleaseDirectory "excel-formula-remover-cli-$Version-win-x64.zip"
    $executableName = 'excel-formula-remover.exe'
}
elseif ($IsLinux) {
    $archive = Join-Path $ReleaseDirectory "excel-formula-remover-cli-$Version-linux-x64.tar.gz"
    $executableName = 'excel-formula-remover'
}
else {
    throw 'No CLI archive is published for this operating system.'
}

function Invoke-Installer {
    param(
        [Parameter(Mandatory)] [string] $ChecksumPath,
        [Parameter(Mandatory)] [string] $BinDirectory
    )

    if ($IsWindows) {
        # Run under the inbox Windows PowerShell 5.1: it is the host a user
        # without PowerShell 7 has, and the stricter of the two.
        $shell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        & $shell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ReleaseDirectory 'install.ps1') `
            -Archive $archive -Checksums $ChecksumPath -BinDirectory $BinDirectory | Write-Host
        if ($LASTEXITCODE -ne 0) {
            throw "install.ps1 failed with exit code $LASTEXITCODE."
        }

        return
    }

    & sh (Join-Path $ReleaseDirectory 'install.sh') $archive $ChecksumPath $BinDirectory | Write-Host
    if ($LASTEXITCODE -ne 0) {
        throw "install.sh failed with exit code $LASTEXITCODE."
    }
}

# Stay beneath the repository: a shell started by a packaged desktop app can
# have the user's temporary directory redirected, which the installer refuses.
$workDirectory = Join-Path $repositoryRoot "artifacts/artifact-test/$([Guid]::NewGuid().ToString('N'))"
$binDirectory = Join-Path $workDirectory 'bin'
New-Item -ItemType Directory -Path $workDirectory | Out-Null

try {
    # Positive control: the installer must reject an archive whose checksum
    # does not match, and must not leave an executable behind.
    $tamperedChecksums = Join-Path $workDirectory 'SHA256SUMS.txt'
    $archiveName = [System.IO.Path]::GetFileName($archive)
    [System.IO.File]::WriteAllLines(
        $tamperedChecksums,
        [string[]] @(('0' * 64) + "  $archiveName"))
    $rejected = $false
    try {
        Invoke-Installer -ChecksumPath $tamperedChecksums -BinDirectory $binDirectory 2>$null
    }
    catch {
        $rejected = $true
    }

    if (-not $rejected -or (Test-Path -LiteralPath (Join-Path $binDirectory $executableName))) {
        throw 'The installer accepted an archive with a mismatched checksum.'
    }

    Invoke-Installer -ChecksumPath $checksums -BinDirectory $binDirectory
    # Reinstalling over an existing executable is the documented upgrade path.
    Invoke-Installer -ChecksumPath $checksums -BinDirectory $binDirectory

    $executable = Join-Path $binDirectory $executableName
    $leftovers = @(Get-ChildItem -LiteralPath $binDirectory -Force | Where-Object Name -NE $executableName)
    if ($leftovers.Count -ne 0) {
        throw "The installer left unexpected entries: $($leftovers.Name -join ', ')"
    }

    $env:EXCEL_FORMULA_REMOVER_EXE = $executable
    $env:EXCEL_FORMULA_REMOVER_EXPECTED_VERSION = $Version
    try {
        Write-Host "> dotnet test (process contract) against $executable"
        & dotnet test $testProject --configuration Release `
            --filter 'FullyQualifiedName~ProcessContractTests' `
            --logger 'trx;LogFileName=artifact.trx' --results-directory $workDirectory
        if ($LASTEXITCODE -ne 0) {
            throw "Process contract tests failed against '$executable'."
        }

        # A filter that matches nothing also exits 0; require executed tests.
        [xml] $testRun = Get-Content -LiteralPath (Join-Path $workDirectory 'artifact.trx') -Raw
        $counters = $testRun.TestRun.ResultSummary.Counters
        if ([int] $counters.executed -lt 1 -or $counters.executed -ne $counters.passed) {
            throw "Expected every process contract test to run and pass; executed $($counters.executed), passed $($counters.passed)."
        }
    }
    finally {
        Remove-Item Env:EXCEL_FORMULA_REMOVER_EXE, Env:EXCEL_FORMULA_REMOVER_EXPECTED_VERSION
    }

    Write-Host "Artifact '$archiveName' passed installation and process contract checks."
}
finally {
    Remove-Item -LiteralPath $workDirectory -Recurse -Force
}
