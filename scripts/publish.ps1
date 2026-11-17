[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [string] $Version
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$semanticVersionPattern = '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-((?:0|[1-9][0-9]*|[0-9]*[A-Za-z-][0-9A-Za-z-]*)(?:\.(?:0|[1-9][0-9]*|[0-9]*[A-Za-z-][0-9A-Za-z-]*))*))?$'

if ($Version.Length -gt 128 -or $Version -cnotmatch $semanticVersionPattern) {
    throw "Version '$Version' is not a valid semantic version (for example, 1.2.3 or 1.2.3-rc.1). Build metadata is not accepted because the executable reports its version without it."
}

$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$guiProject = Join-Path $repositoryRoot 'src/ExcelFormulaRemover.App/ExcelFormulaRemover.App.csproj'
$cliProject = Join-Path $repositoryRoot 'src/ExcelFormulaRemover.Cli/ExcelFormulaRemover.Cli.csproj'
$publishRoot = [System.IO.Path]::GetFullPath((Join-Path $repositoryRoot 'artifacts/publish'))
$versionDirectory = [System.IO.Path]::GetFullPath((Join-Path $publishRoot $Version))
$pathSeparators = [char[]] @(
    [System.IO.Path]::DirectorySeparatorChar,
    [System.IO.Path]::AltDirectorySeparatorChar
)
$publishPrefix = $publishRoot.TrimEnd($pathSeparators) + [System.IO.Path]::DirectorySeparatorChar

if (-not $versionDirectory.StartsWith($publishPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "The publish destination must remain beneath '$publishRoot'."
}

if (-not (Test-Path -LiteralPath $guiProject -PathType Leaf)) {
    throw "GUI project not found at '$guiProject'."
}

if (-not (Test-Path -LiteralPath $cliProject -PathType Leaf)) {
    throw "CLI project not found at '$cliProject'."
}

if (Test-Path -LiteralPath $versionDirectory) {
    throw "Release '$Version' already exists at '$versionDirectory'. Versioned releases are immutable."
}

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

New-Item -ItemType Directory -Path $publishRoot -Force | Out-Null

$stagingDirectory = Join-Path $publishRoot ".staging-$([Guid]::NewGuid().ToString('N'))"
$guiOutput = Join-Path $stagingDirectory 'gui'
$bundleDirectory = Join-Path $stagingDirectory 'release'
$installDirectory = Join-Path $repositoryRoot 'scripts/install'

# The CLI targets this release claims. Each is certified by test-artifact.ps1
# on its own operating system before the release is published.
$cliTargets = @(
    @{ Runtime = 'win-x64'; Executable = 'excel-formula-remover.exe'; Extension = 'zip' }
    @{ Runtime = 'linux-x64'; Executable = 'excel-formula-remover'; Extension = 'tar.gz' }
)

New-Item -ItemType Directory -Path $guiOutput, $bundleDirectory | Out-Null

# A runtime-specific publish needs a RID-specific assets graph. Point NuGet at
# an unused staging lock path so the restore cannot rewrite the repository's
# neutral locks, which check.ps1 verifies independently.
function Restore-ForRuntime {
    param(
        [Parameter(Mandatory)] [string] $Project,
        [Parameter(Mandatory)] [string] $Runtime
    )

    $lockPath = Join-Path $stagingDirectory "rid-restore/$Runtime/$([Guid]::NewGuid().ToString('N'))/packages.lock.json"
    Invoke-NativeCommand -FilePath 'dotnet' -ArgumentList @(
        'restore',
        $Project,
        '--runtime',
        $Runtime,
        '-p:RestoreLockedMode=false',
        '-p:RestorePackagesWithLockFile=false',
        "-p:NuGetLockFilePath=$lockPath"
    )
}

Push-Location $repositoryRoot
try {
    $archives = @()

    foreach ($target in $cliTargets) {
        $runtime = $target.Runtime
        $cliOutput = Join-Path $stagingDirectory "cli-$runtime"
        Restore-ForRuntime -Project $cliProject -Runtime $runtime

        # Untrimmed on purpose: the OpenXML SDK relies on reflection that
        # trimming and Native AOT do not preserve.
        Invoke-NativeCommand -FilePath 'dotnet' -ArgumentList @(
            'publish',
            $cliProject,
            '--configuration',
            'Release',
            '--runtime',
            $runtime,
            '--self-contained',
            'true',
            '--no-restore',
            '--output',
            $cliOutput,
            '-p:PublishSingleFile=true',
            '-p:PublishTrimmed=false',
            '-p:EnableCompressionInSingleFile=true',
            '-p:IncludeNativeLibrariesForSelfExtract=true',
            '-p:DebugSymbols=false',
            '-p:DebugType=None',
            '-p:ContinuousIntegrationBuild=true',
            "-p:Version=$Version",
            "-p:InformationalVersion=$Version"
        )

        # The installers accept an archive holding exactly the executable.
        $published = @(Get-ChildItem -LiteralPath $cliOutput -Force)
        if ($published.Count -ne 1 -or $published[0].Name -cne $target.Executable) {
            throw "The $runtime CLI publish produced '$($published.Name -join ', ')' instead of only '$($target.Executable)'."
        }

        $archive = Join-Path $bundleDirectory "excel-formula-remover-cli-$Version-$runtime.$($target.Extension)"
        if ($target.Extension -eq 'zip') {
            Compress-Archive -LiteralPath $published[0].FullName -DestinationPath $archive -CompressionLevel Optimal
        }
        else {
            # Use the Windows-provided tar: another tar earlier in PATH, such
            # as Git's, can misread drive-letter paths.
            $tar = Join-Path $env:SystemRoot 'System32\tar.exe'
            Invoke-NativeCommand -FilePath $tar -ArgumentList @(
                '-czf', $archive, '-C', $cliOutput, $target.Executable
            )
        }

        $archives += $archive
    }

    Restore-ForRuntime -Project $guiProject -Runtime 'win-x64'

    Invoke-NativeCommand -FilePath 'dotnet' -ArgumentList @(
        'publish',
        $guiProject,
        '--configuration',
        'Release',
        '--runtime',
        'win-x64',
        '--self-contained',
        'true',
        '--no-restore',
        '--output',
        $guiOutput,
        '-p:PublishSingleFile=true',
        '-p:PublishTrimmed=false',
        '-p:IncludeNativeLibrariesForSelfExtract=true',
        '-p:DebugSymbols=false',
        '-p:DebugType=None',
        '-p:ContinuousIntegrationBuild=true',
        "-p:Version=$Version",
        "-p:InformationalVersion=$Version"
    )

    if (-not (Get-ChildItem -LiteralPath $guiOutput -File)) {
        throw 'The GUI publish completed without producing files.'
    }

    $guiArchive = Join-Path $bundleDirectory "excel-formula-remover-gui-$Version-win-x64.zip"
    Compress-Archive -Path (Join-Path $guiOutput '*') -DestinationPath $guiArchive -CompressionLevel Optimal
    $archives += $guiArchive

    $checksumLines = foreach ($archivePath in ($archives | Sort-Object)) {
        $hash = Get-FileHash -LiteralPath $archivePath -Algorithm SHA256
        "$($hash.Hash.ToLowerInvariant())  $([System.IO.Path]::GetFileName($archivePath))"
    }

    $checksumPath = Join-Path $bundleDirectory 'SHA256SUMS.txt'
    [System.IO.File]::WriteAllLines(
        $checksumPath,
        $checksumLines,
        [System.Text.UTF8Encoding]::new($false)
    )

    # Ship the installers with the release so a consumer needs no checkout.
    foreach ($installerName in 'install.ps1', 'install.sh', 'verify-windows-install.ps1') {
        Copy-Item -LiteralPath (Join-Path $installDirectory $installerName) -Destination $bundleDirectory
    }

    # Certify the Windows archive before it becomes a versioned release. The
    # Linux archive is certified by the same script on a Linux host.
    & (Join-Path $PSScriptRoot 'test-artifact.ps1') -ReleaseDirectory $bundleDirectory -Version $Version

    [System.IO.Directory]::Move($bundleDirectory, $versionDirectory)

    Write-Host "Release artifacts created at '$versionDirectory':"
    Get-ChildItem -LiteralPath $versionDirectory -File |
        Sort-Object -Property Name |
        ForEach-Object { Write-Host "  $($_.Name)" }
}
finally {
    Pop-Location

    if (Test-Path -LiteralPath $stagingDirectory) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force
    }
}
