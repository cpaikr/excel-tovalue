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

# A consumer copies one skill directory out of the repository, so every
# resource it links must live inside that directory.
function Test-SkillPackages {
    $skillRoot = Join-Path $repositoryRoot 'skill'
    $problems = [System.Collections.Generic.List[string]]::new()

    foreach ($package in Get-ChildItem -LiteralPath $skillRoot -Directory) {
        $entryPoint = Join-Path $package.FullName 'SKILL.md'
        if (-not (Test-Path -LiteralPath $entryPoint -PathType Leaf)) {
            $problems.Add("$($package.Name): SKILL.md is missing.")
            continue
        }

        $entryText = Get-Content -LiteralPath $entryPoint -Raw
        $frontmatter = [regex]::Match($entryText, '\A---\r?\n(?<body>.*?)\r?\n---\r?\n', 'Singleline')
        $name = [regex]::Match($frontmatter.Groups['body'].Value, '(?m)^name:\s*(?<value>\S+)\s*$')
        if (-not $frontmatter.Success -or $name.Groups['value'].Value -cne $package.Name) {
            $problems.Add("$($package.Name): frontmatter 'name' must equal the directory name.")
        }

        if ($frontmatter.Groups['body'].Value -notmatch '(?m)^description:\s*\S') {
            $problems.Add("$($package.Name): frontmatter 'description' is missing.")
        }

        $packagePrefix = $package.FullName + [System.IO.Path]::DirectorySeparatorChar
        foreach ($document in Get-ChildItem -LiteralPath $package.FullName -Recurse -File -Filter '*.md') {
            $documentText = Get-Content -LiteralPath $document.FullName -Raw
            foreach ($link in [regex]::Matches($documentText, '\]\((?<target>[^)\s]+)\)')) {
                $target = $link.Groups['target'].Value
                if ($target -match '^(https?:|mailto:|#)') {
                    continue
                }

                $relativePath = ($target -split '#', 2)[0]
                $resolved = [System.IO.Path]::GetFullPath((Join-Path $document.DirectoryName $relativePath))
                if (-not $resolved.StartsWith($packagePrefix, [StringComparison]::Ordinal)) {
                    $problems.Add("$($document.Name): link '$target' leaves the skill package.")
                }
                elseif (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
                    $problems.Add("$($document.Name): link '$target' does not resolve.")
                }
            }
        }
    }

    if ($problems.Count -gt 0) {
        throw "Skill package check failed:`n  $($problems -join "`n  ")"
    }

    Write-Host 'Skill packages are self-contained.'
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

    Test-SkillPackages

    Write-Host 'Repository checks completed successfully.'
}
finally {
    Pop-Location
}
