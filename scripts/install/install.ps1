# Install the xlclean CLI on Windows x64.
# Use this installer, the archive, and SHA256SUMS.txt from the same release.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $Archive,
    [Parameter(Mandatory)][string] $Checksums,
    [string] $BinDirectory = (Join-Path $env:LOCALAPPDATA 'xlclean\bin')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$executableName = 'xlclean.exe'
$archivePattern = '^xlclean-cli-(?<version>[0-9A-Za-z.-]+)-win-x64\.zip$'

if ([Environment]::OSVersion.Platform -ne 'Win32NT' -or
    [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -ne 'X64') {
    throw 'This installer requires Windows x64.'
}

if (-not ('XlCleanInstallerNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

public static class XlCleanInstallerNative
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandle(
        SafeFileHandle hFile,
        StringBuilder lpszFilePath,
        uint cchFilePath,
        uint dwFlags);

    public static string GetFinalPath(string path)
    {
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete))
        {
            var capacity = 512;
            while (true)
            {
                var buffer = new StringBuilder(capacity);
                var length = GetFinalPathNameByHandle(stream.SafeFileHandle, buffer,
                    (uint)buffer.Capacity, 0);
                if (length == 0)
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(),
                        "GetFinalPathNameByHandle failed.");
                }
                if (length < buffer.Capacity)
                {
                    return buffer.ToString();
                }
                capacity = checked((int)length + 1);
            }
        }
    }
}
'@
}

function ConvertTo-ComparableWindowsPath([string] $Path) {
    if ($Path.StartsWith('\\?\UNC\', [StringComparison]::OrdinalIgnoreCase)) {
        return '\\' + $Path.Substring(8)
    }
    if ($Path.StartsWith('\\?\', [StringComparison]::OrdinalIgnoreCase)) {
        return $Path.Substring(4)
    }
    return [IO.Path]::GetFullPath($Path)
}

# A shell started by a packaged (MSIX) desktop app, including an agent's shell,
# can have writes under AppData redirected to private storage that no other
# program sees. Installing there would report success for an executable the
# user's own terminal cannot find, so refuse it before changing anything.
function Assert-InstallationPathIsVisible([string] $Path) {
    $probe = Join-Path $Path ('.xlclean-visibility-' + [Guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($probe, 'xlclean installer visibility probe')
        try {
            $advertised = ConvertTo-ComparableWindowsPath $probe
            $physical = ConvertTo-ComparableWindowsPath (
                [XlCleanInstallerNative]::GetFinalPath($probe))
        }
        catch {
            throw "Could not verify the physical installation path for '$Path': $($_.Exception.Message) Run this installer from an independently launched PowerShell session, or choose a local directory with -BinDirectory."
        }

        if (-not [string]::Equals($advertised, $physical, [StringComparison]::OrdinalIgnoreCase)) {
            throw @"
The selected installation directory is redirected and resolves to a different physical path.
Advertised path: $advertised
Physical path:   $physical
No existing installation was changed. Run this installer from an independently launched PowerShell session, or rerun it with -BinDirectory "$env:USERPROFILE\.local\bin" (or another non-redirected directory).
"@
        }
    }
    finally {
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    }
}

# ProviderPath, not Path: on a UNC location Path carries a provider prefix
# that .NET file APIs reject.
$Archive = (Resolve-Path -LiteralPath $Archive).ProviderPath
$archiveName = [IO.Path]::GetFileName($Archive)
if ($archiveName -cnotmatch $archivePattern) {
    throw "Expected an archive named xlclean-cli-<version>-win-x64.zip, not '$archiveName'."
}
$version = $Matches['version']

$checksumLines = @(Get-Content -LiteralPath $Checksums |
        Where-Object { $_ -cmatch ('^[0-9a-f]{64}  ' + [regex]::Escape($archiveName) + '$') })
if ($checksumLines.Count -ne 1) {
    throw "Missing or duplicate checksum for '$archiveName'."
}
$expectedHash = $checksumLines[0].Substring(0, 64)
if ((Get-FileHash -LiteralPath $Archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expectedHash) {
    throw 'Archive checksum mismatch; the existing installation was not changed.'
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($Archive)
try {
    $entryNames = @($zip.Entries | ForEach-Object { $_.FullName })
}
finally {
    $zip.Dispose()
}
# The archive carries the license terms and third-party notices beside the
# executable. Only the executable is installed, so a shared bin directory
# gains one file.
if (($entryNames -join '|') -cne "$executableName|LICENSE.md|THIRD-PARTY-NOTICES.txt") {
    throw "Unexpected archive contents; expected only $executableName, LICENSE.md, and THIRD-PARTY-NOTICES.txt."
}

$null = New-Item -ItemType Directory -Path $BinDirectory -Force
$BinDirectory = (Resolve-Path -LiteralPath $BinDirectory).ProviderPath
Assert-InstallationPathIsVisible $BinDirectory

$stage = Join-Path $BinDirectory ('.xlclean-install-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $stage
try {
    Expand-Archive -LiteralPath $Archive -DestinationPath $stage
    $candidate = Join-Path $stage $executableName

    # Check the candidate before it replaces a working installation.
    $reported = (& $candidate --version) -join ' '
    if ($LASTEXITCODE -ne 0 -or $reported -cne "xlclean $version") {
        throw "The candidate reported '$reported' instead of 'xlclean $version'; the existing installation was not changed."
    }

    $destination = Join-Path $BinDirectory $executableName
    if (Test-Path -LiteralPath $destination) {
        # PowerShell converts $null to an empty string for .NET string
        # parameters, which File.Replace rejects as a backup path.
        [IO.File]::Replace($candidate, $destination, [NullString]::Value)
    }
    else {
        [IO.File]::Move($candidate, $destination)
    }

    Write-Output "Installed xlclean $version to $destination"
    Write-Output "Add $BinDirectory to the user PATH if needed; new terminals pick up a persistent PATH change."
}
finally {
    Remove-Item -LiteralPath $stage -Recurse -Force
}
