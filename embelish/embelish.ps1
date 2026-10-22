#requires -Version 5.1

# .EXTERNALHELP ../docs/embelish/embelish-help.xml

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium', PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Embelish', 'Deembelish')]
    [string] $Mode,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $SourcePath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $DestinationPath,

    [Parameter()]
    [ValidateRange(1, 1048576)]
    [int] $VolumeSizeMegabytes = 10,

    [Parameter()]
    [ValidateRange(0, 9)]
    [int] $CompressionLevel = 9,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $SevenZipPath,

    [Parameter()]
    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string] $ArchiveBaseName = 'embelish',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $LogDirectory,

    [Parameter()]
    [switch] $Force
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$script:ToolName = 'embelish'
$script:ToolVersion = '1.2.2'
$script:RunId = [guid]::NewGuid().ToString('N').Substring(0, 12)
$script:LogPath = $null
$script:DebugEnabled = $false
$script:LastSevenZipExit = $null
$script:OperationResult = $null

function Get-EmbelishTimestamp {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    [OutputType([string])]
    param()

    return [datetime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Write-DiagnosticEvent {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Stage,

        [Parameter()]
        [ValidateSet('info', 'warn', 'error')]
        [string] $Level = 'info',

        [Parameter()]
        [hashtable] $Data
    )

    if ($script:DebugEnabled) {
        Write-Debug ("embelish diagnostic level={0}" -f $Level)
    }
    if ([string]::IsNullOrEmpty($script:LogPath)) {
        return
    }

    $record = [ordered]@{
        ts    = Get-EmbelishTimestamp
        runId = $script:RunId
        tool  = $script:ToolName
        ver   = $script:ToolVersion
        level = $Level
        stage = $(if ($Stage -cmatch '^(run|embelish|deembelish)\.(start|end|error|encrypted|extracted|done)$') { $Stage } else { 'run.error' })
    }
    if ($PSBoundParameters.ContainsKey('Data') -and $null -ne $Data) {
        foreach ($key in @('volumeMb', 'mx', 'exit', 'exitCode', 'shards', 'restored', 'line')) {
            if ($Data.ContainsKey($key) -and $Data[$key] -is [int]) {
                $record[$key] = $Data[$key]
            }
        }
        if ($Data.ContainsKey('errorKind') -and $Data.errorKind -cin @('Input', 'Operational', 'Unexpected')) {
            $record['errorKind'] = $Data.errorKind
        }
        if ($Data.ContainsKey('hresult') -and $Data.hresult -is [string] -and $Data.hresult -cmatch '\A0x[0-9A-F]{8}\z') {
            $record['hresult'] = $Data.hresult
        }
        foreach ($key in @('succeeded', 'whatIf')) {
            if ($Data.ContainsKey($key) -and $Data[$key] -is [bool]) {
                $record[$key] = $Data[$key]
            }
        }
    }

    try {
        $line = ($record | ConvertTo-Json -Depth 6 -Compress) + "`n"
        [System.IO.File]::AppendAllText($script:LogPath, $line, [System.Text.Encoding]::UTF8)
    } catch {
        Write-Warning 'embelish: unable to write a diagnostic event.'
    }
}

function Initialize-DiagnosticContext {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    param(
        [Parameter()]
        [string] $RequestedDirectory
    )

    if (-not [string]::IsNullOrEmpty($RequestedDirectory)) {
        $directory = $RequestedDirectory
    } else {
        $localAppData = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
        if ([string]::IsNullOrEmpty($localAppData)) {
            $localAppData = [System.IO.Path]::GetTempPath()
        }
        $directory = [System.IO.Path]::Combine($localAppData, 'Embelish', 'logs')
    }

    try {
        $directory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($directory)
        if (-not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        $stamp = [datetime]::UtcNow.ToString('yyyyMMddTHHmmssZ', [System.Globalization.CultureInfo]::InvariantCulture)
        $fileName = "{0}-{1}-{2}.jsonl" -f $script:ToolName, $stamp, $script:RunId
        $script:LogPath = [System.IO.Path]::Combine($directory, $fileName)

        $existing = Get-ChildItem -LiteralPath $directory -Filter 'embelish-*.jsonl' -File -ErrorAction SilentlyContinue |
            Sort-Object -Property LastWriteTimeUtc -Descending | Select-Object -Skip 30
        foreach ($stale in $existing) {
            Remove-Item -LiteralPath $stale.FullName -Force -ErrorAction SilentlyContinue
        }
    } catch {
        $script:LogPath = $null
        Write-Warning 'embelish: diagnostics disabled; the log directory could not be prepared.'
    }
}

function Resolve-SevenZipPath {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()]
        [string] $ExplicitPath
    )

    if (-not [string]::IsNullOrEmpty($ExplicitPath)) {
        if (Test-Path -LiteralPath $ExplicitPath -PathType Leaf) {
            return (Resolve-Path -LiteralPath $ExplicitPath).Path
        }
        throw [System.IO.FileNotFoundException]::new("7-Zip executable not found at supplied -SevenZipPath '$ExplicitPath'.")
    }

    foreach ($candidate in @('7z', '7z.exe', '7za', '7za.exe', '7zz', '7zz.exe', '7zr', '7zr.exe')) {
        $command = Get-Command -Name $candidate -CommandType Application -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($null -ne $command) {
            return $command.Source
        }
    }

    foreach ($fixed in @(
            [System.IO.Path]::Combine($env:ProgramFiles, '7-Zip', '7z.exe'),
            [System.IO.Path]::Combine(${env:ProgramFiles(x86)}, '7-Zip', '7z.exe'))) {
        if (-not [string]::IsNullOrEmpty($fixed) -and (Test-Path -LiteralPath $fixed -PathType Leaf)) {
            return $fixed
        }
    }

    throw [System.IO.FileNotFoundException]::new('7-Zip executable not found. Install 7-Zip or pass -SevenZipPath.')
}

function Invoke-SevenZipCommand {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Executable,

        [Parameter(Mandatory = $true)]
        [string[]] $ArgumentList
    )

    $ErrorActionPreference = 'Stop'
    if (Test-Path -Path 'variable:PSNativeCommandUseErrorActionPreference') {
        Set-Variable -Name 'PSNativeCommandUseErrorActionPreference' -Value $false -Scope Local
    }

    $script:LastSevenZipExit = $null
    & $Executable @ArgumentList
    if ($null -eq $LASTEXITCODE) {
        throw [System.InvalidOperationException]::new('7-Zip did not return a native exit code.')
    }
    $script:LastSevenZipExit = [int]$LASTEXITCODE
}

function Resolve-SevenZipOutcome {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [int] $ExitCode
    )

    switch ($ExitCode) {
        0 { $status = 'Success'; $category = 'ok' }
        1 { $status = 'CompletedWithWarning'; $category = 'warning' }
        2 { $status = 'FatalError'; $category = 'operational' }
        7 { $status = 'CommandLineError'; $category = 'input' }
        8 { $status = 'NotEnoughMemory'; $category = 'operational' }
        255 { $status = 'StoppedOrWrongPassword'; $category = 'operational' }
        default { $status = 'UnknownFailure'; $category = 'unknown' }
    }

    return [pscustomobject]@{
        ExitCode  = $ExitCode
        Status    = $status
        Category  = $category
        IsSuccess = ($category -eq 'ok')
        IsWarning = ($category -eq 'warning')
    }
}

function Assert-SevenZipOutcome {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject] $Outcome
    )

    if ($Outcome.IsSuccess) {
        return
    }

    $message = "7-Zip failed with exit code $($Outcome.ExitCode) ($($Outcome.Status))."
    if ($Outcome.Category -eq 'input') {
        throw [System.ArgumentException]::new($message)
    }
    if ($Outcome.Category -eq 'operational' -or $Outcome.IsWarning) {
        throw [System.InvalidOperationException]::new($message)
    }
    throw [System.Exception]::new($message)
}

function New-CleanOutputDirectory {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter()]
        [switch] $AllowExisting
    )

    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        throw [System.ArgumentException]::new("Destination '$fullPath' is a file; a directory path is required.")
    }

    if (Test-Path -LiteralPath $fullPath -PathType Container) {
        $hasChildren = @(Get-ChildItem -LiteralPath $fullPath -Force -ErrorAction SilentlyContinue).Count -gt 0
        if ($hasChildren -and -not $AllowExisting) {
            throw [System.ArgumentException]::new("Destination '$fullPath' already exists and is not empty; pass -Force to overwrite.")
        }
    } elseif ($PSCmdlet.ShouldProcess($fullPath, 'Create destination directory')) {
        New-Item -ItemType Directory -Path $fullPath -Force | Out-Null
    }

    return $fullPath
}

function ConvertTo-Base64Shard {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $VolumePath,

        [Parameter(Mandatory = $true)]
        [string] $ShardPath
    )

    Copy-Base64Stream -InputPath $VolumePath -OutputPath $ShardPath -Encode $true
}

function ConvertFrom-Base64Shard {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    param([string] $ShardPath, [string] $VolumePath)
    Copy-Base64Stream -InputPath $ShardPath -OutputPath $VolumePath -Encode $false
}

function Copy-Base64Stream {
    [CmdletBinding()]
    param([string] $InputPath, [string] $OutputPath, [bool] $Encode)
    $inputStream = [System.IO.File]::OpenRead($InputPath)
    $outputStream = $null
    $crypto = $null
    $transform = $null
    try {
        $outputStream = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        if (-not $Encode) {
            Read-Base64Stream -InputStream $inputStream -OutputStream $outputStream
            return
        }
        $transform = [System.Security.Cryptography.ToBase64Transform]::new()
        $crypto = [System.Security.Cryptography.CryptoStream]::new($outputStream, $transform,
            [System.Security.Cryptography.CryptoStreamMode]::Write)
        $inputStream.CopyTo($crypto, 81920)
        $crypto.FlushFinalBlock()
    } finally {
        if ($null -ne $crypto) { $crypto.Dispose() }
        if ($null -ne $outputStream) { $outputStream.Dispose() }
        if ($null -ne $transform) { $transform.Dispose() }
        $inputStream.Dispose()
    }
}

function Read-Base64Stream {
    [CmdletBinding()]
    param([System.IO.Stream] $InputStream, [System.IO.Stream] $OutputStream)
    $reader = [System.IO.StreamReader]::new($InputStream, [System.Text.Encoding]::ASCII, $false, 8192, $true)
    $buffer = [char[]]::new(8192)
    $carry = ''
    $padded = $false
    try {
        while (($count = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $text = [regex]::Replace([string]::new($buffer, 0, $count), '[ \t\r\n]', '')
            if ($padded -and $text.Length -gt 0) {
                throw [System.FormatException]::new('Base64 data follows the final padded group.')
            }
            $text = $carry + $text
            $complete = $text.Length - ($text.Length % 4)
            $carry = $text.Substring($complete)
            if ($complete -gt 0) {
                $fragment = $text.Substring(0, $complete)
                $bytes = [System.Convert]::FromBase64String($fragment)
                $OutputStream.Write($bytes, 0, $bytes.Length)
                $padded = $fragment.IndexOf('=') -ge 0
            }
        }
        if ($carry.Length -ne 0) {
            throw [System.FormatException]::new('The final base64 group is incomplete.')
        }
    } finally {
        $reader.Dispose()
    }
}

function Resolve-ShardManifest {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory = $true)][string] $Path)
    $pattern = '^(?<base>.+\.7z)\.(?<num>[0-9]+)\.txt$'
    $selectedBase = $null
    if (-not (Test-Path -LiteralPath $Path)) {
        throw [System.ArgumentException]::new('The selected shard source does not exist.')
    }
    $item = Get-Item -LiteralPath $Path -Force
    if ($item.PSIsContainer) {
        $directory = $item.FullName
    } else {
        $match = [regex]::Match($item.Name, $pattern)
        if (-not $match.Success) {
            throw [System.ArgumentException]::new('Select an archive shard file or its containing directory.')
        }
        $selectedBase = $match.Groups['base'].Value
        $directory = $item.DirectoryName
    }
    $shards = @(Get-ChildItem -LiteralPath $directory -File -Force | Where-Object {
        $_.Name -cmatch $pattern -and ($null -eq $selectedBase -or
            [regex]::Match($_.Name, $pattern).Groups['base'].Value -ceq $selectedBase)
    })
    $bases = @($shards | ForEach-Object {
        [regex]::Match($_.Name, $pattern).Groups['base'].Value
    } | Sort-Object -Unique)
    if ($bases.Count -ne 1 -or $shards.Count -gt 10000) {
        throw [System.ArgumentException]::new('Select exactly one shard set, with 1 through 10000 volumes; select a shard file when several sets share a directory.')
    }
    $ordered = @(Get-OrderedVolume -Files $shards -BaseName $bases[0] -Suffix '.txt')
    return [pscustomobject]@{ Directory = $directory; BaseName = $bases[0]; Shards = $ordered }
}

function Get-OrderedVolume {
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param([object[]] $Files, [string] $BaseName, [string] $Suffix = '')
    if ($Files.Count -lt 1 -or $Files.Count -gt 10000) {
        throw [System.ArgumentException]::new('A nonempty bounded volume set is required.')
    }
    $numbered = @{}
    foreach ($file in $Files) {
        $number = 0
        $pattern = '^' + [regex]::Escape($BaseName) + '\.(?<num>[0-9]+)' + [regex]::Escape($Suffix) + '$'
        $match = [regex]::Match($file.Name, $pattern)
        if (-not $match.Success -or -not [int]::TryParse($match.Groups['num'].Value, [ref]$number)) {
            throw [System.ArgumentException]::new('A volume filename is invalid.')
        }
        if ($number -lt 1 -or $number -gt 10000 -or $match.Groups['num'].Value -cne $number.ToString('D3') -or
            $numbered.ContainsKey($number) -or $file.Length -le 0) {
            throw [System.ArgumentException]::new('Volumes must be nonempty, uniquely and canonically numbered from 001.')
        }
        $numbered[$number] = $file
    }
    for ($index = 1; $index -le $Files.Count; $index++) {
        if (-not $numbered.ContainsKey($index)) {
            throw [System.ArgumentException]::new('The volume set has a numbering gap.')
        }
        $numbered[$index]
    }
}

function Assert-EmbelishPath {
    [CmdletBinding()]
    param([string] $Path)
    $current = $Path
    while (-not [string]::IsNullOrEmpty($current)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if ($item.LinkType -in @('SymbolicLink', 'Junction', 'HardLink')) {
                throw [System.ArgumentException]::new('Use actual filesystem paths: linked ancestry cannot protect publication or detect source/output overlap.')
            }
        }
        $current = [System.IO.Path]::GetDirectoryName($current)
    }
}

function Get-EmbelishPathIdentity {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Path, [switch] $ForFileSystem)
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
        return $Path
    }
    if (-not ('ConsultHelp.EmbelishPathV1' -as [type])) {
        Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace ConsultHelp {
    public static class EmbelishPathV1 {
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern SafeFileHandle CreateFile(string name, uint access, uint share,
            IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern uint GetFinalPathNameByHandle(SafeFileHandle handle, StringBuilder path, uint length, uint flags);
        [StructLayout(LayoutKind.Sequential)] struct FILE_ID_INFO {
            public ulong VolumeSerialNumber, FileIdLow, FileIdHigh;
        }
        [DllImport("kernel32.dll", SetLastError=true)]
        static extern bool GetFileInformationByHandleEx(SafeFileHandle handle, int kind, out FILE_ID_INFO info, uint size);
        static FILE_ID_INFO FileId(string path) {
            using (SafeFileHandle handle = CreateFile(path, 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero)) {
                if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
                FILE_ID_INFO id;
                if (!GetFileInformationByHandleEx(handle, 18, out id, (uint)Marshal.SizeOf(typeof(FILE_ID_INFO))))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                return id;
            }
        }
        static bool ContainsIdentity(string path, FILE_ID_INFO expected) {
            while (!String.IsNullOrEmpty(path)) {
                if (System.IO.Directory.Exists(path)) {
                    FILE_ID_INFO current = FileId(path);
                    if (current.VolumeSerialNumber == expected.VolumeSerialNumber &&
                        current.FileIdLow == expected.FileIdLow && current.FileIdHigh == expected.FileIdHigh) return true;
                }
                path = System.IO.Path.GetDirectoryName(path);
            }
            return false;
        }
        public static bool Overlaps(string source, string destination) {
            if (ContainsIdentity(destination, FileId(source))) return true;
            return System.IO.Directory.Exists(destination) && ContainsIdentity(source, FileId(destination));
        }
        public static string Identity(string path, bool forFileSystem) {
            using (SafeFileHandle handle = CreateFile(path, 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero)) {
                if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
                StringBuilder name = new StringBuilder(512);
                uint flags = forFileSystem ? 0u : 2u;
                uint length = GetFinalPathNameByHandle(handle, name, (uint)name.Capacity, flags);
                if (length == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
                if (length >= name.Capacity) {
                    name = new StringBuilder(checked((int)length + 1));
                    length = GetFinalPathNameByHandle(handle, name, (uint)name.Capacity, flags);
                    if (length == 0 || length >= name.Capacity) throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                string result = name.ToString();
                if (forFileSystem && result.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
                    return @"\\" + result.Substring(8);
                if (forFileSystem && result.StartsWith(@"\\?\", StringComparison.Ordinal))
                    return result.Substring(4);
                return result;
            }
        }
    }
}
'@
    }
    $suffix = [System.Collections.Generic.List[string]]::new()
    $ancestor = $Path
    while (-not [System.IO.Directory]::Exists($ancestor)) {
        $suffix.Insert(0, [System.IO.Path]::GetFileName($ancestor))
        $parent = [System.IO.Path]::GetDirectoryName($ancestor)
        if ([string]::IsNullOrEmpty($parent)) {
            throw [System.ArgumentException]::new('The path has no existing directory ancestor.')
        }
        $ancestor = $parent
    }
    $identity = [ConsultHelp.EmbelishPathV1]::Identity($ancestor, [bool]$ForFileSystem)
    foreach ($segment in $suffix) { $identity = [System.IO.Path]::Combine($identity, $segment) }
    return $identity.TrimEnd([System.IO.Path]::DirectorySeparatorChar)
}

function Get-EmbelishDestination {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Source, [string] $Destination, [bool] $AllowExisting)
    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destination)
    $fullPath = [System.IO.Path]::GetFullPath($fullPath)
    if ($fullPath -eq [System.IO.Path]::GetPathRoot($fullPath)) {
        throw [System.ArgumentException]::new('The destination must be a non-root directory path.')
    }
    $fullPath = $fullPath.TrimEnd([System.IO.Path]::DirectorySeparatorChar)
    Assert-EmbelishPath -Path $Source
    Assert-EmbelishPath -Path $fullPath
    $fullPath = Get-EmbelishPathIdentity -Path $fullPath -ForFileSystem
    $comparison = [System.StringComparison]::Ordinal
    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
        $comparison = [System.StringComparison]::OrdinalIgnoreCase
    }
    $separator = [System.IO.Path]::DirectorySeparatorChar
    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
        if ([ConsultHelp.EmbelishPathV1]::Overlaps($Source, $fullPath)) {
            throw [System.ArgumentException]::new('Source and destination directories must not overlap through filesystem aliases.')
        }
    }
    $sourceIdentity = Get-EmbelishPathIdentity -Path $Source
    $destinationIdentity = Get-EmbelishPathIdentity -Path $fullPath
    if ($destinationIdentity.Equals($sourceIdentity, $comparison) -or
        $destinationIdentity.StartsWith($sourceIdentity.TrimEnd($separator) + $separator, $comparison) -or
        $sourceIdentity.StartsWith($destinationIdentity + $separator, $comparison)) {
        throw [System.ArgumentException]::new('Source and destination directories must not overlap.')
    }
    if ([System.IO.File]::Exists($fullPath) -or [string]::IsNullOrEmpty([System.IO.Path]::GetFileName($fullPath))) {
        throw [System.ArgumentException]::new('The destination must be a non-root directory path.')
    }
    if ([System.IO.Directory]::Exists($fullPath) -and -not $AllowExisting) {
        if (@(Get-ChildItem -LiteralPath $fullPath -Force).Count -gt 0) {
            throw [System.ArgumentException]::new('A nonempty destination requires -Force.')
        }
    }
    return $fullPath
}

function Get-OutputInventory {
    [CmdletBinding()]
    [OutputType([object[]])]
    param([string] $Path, [string] $Relative = '')
    foreach ($item in @(Get-ChildItem -LiteralPath $Path -Force | Sort-Object -Property Name)) {
        if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            throw [System.IO.IOException]::new('Linked output entries cannot be safely copied or published.')
        }
        $name = [System.IO.Path]::Combine($Relative, $item.Name)
        if ($item.PSIsContainer) {
            [pscustomobject]@{ Name = $name; Kind = 'Directory'; Hash = ''; Length = 0 }
            Get-OutputInventory -Path $item.FullName -Relative $name
        } else {
            if ($item.LinkType -eq 'HardLink') {
                throw [System.IO.IOException]::new('Hard-linked output entries cannot be safely replaced.')
            }
            $stream = [System.IO.File]::OpenRead($item.FullName)
            $hash = [System.Security.Cryptography.SHA256]::Create()
            try { $digest = [System.BitConverter]::ToString($hash.ComputeHash($stream)) } finally {
                $hash.Dispose()
                $stream.Dispose()
            }
            [pscustomobject]@{ Name = $name; Kind = 'File'; Hash = $digest; Length = $item.Length }
        }
    }
}

function Set-OutputAccess {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private protection adapter runs only inside a ShouldProcess-approved output transaction; private staging and old access restrictions are mandatory.') ]
    [CmdletBinding()]
    param([string] $Path, [string] $Source)
    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
        $item = Get-Item -LiteralPath $Path -Force
        if ($item.PSIsContainer) {
            $security = [System.Security.AccessControl.DirectorySecurity]::new()
        } else {
            $security = [System.Security.AccessControl.FileSecurity]::new()
        }
        if ([string]::IsNullOrEmpty($Source)) {
            $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $rule = [System.Security.AccessControl.FileSystemAccessRule]::new($sid, 'FullControl',
                'ContainerInherit, ObjectInherit', 'None', 'Allow')
            $security.AddAccessRule($rule)
            $security.SetAccessRuleProtection($true, $false)
        } else {
            $original = Get-Acl -LiteralPath $Source
            $security.SetSecurityDescriptorSddlForm($original.GetSecurityDescriptorSddlForm('Access'), 'Access')
            $security.SetAccessRuleProtection($true, $true)
        }
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            if ($item.PSIsContainer) {
                [System.IO.Directory]::SetAccessControl($item.FullName, $security)
            } else {
                [System.IO.File]::SetAccessControl($item.FullName, $security)
            }
        } else {
            $extensions = 'System.IO.FileSystemAclExtensions' -as [type]
            if ($null -eq $extensions) {
                throw [System.InvalidOperationException]::new('The current Windows runtime cannot apply filesystem access rules.')
            }
            $itemObject = $item.PSObject.BaseObject
            $securityObject = $security.PSObject.BaseObject
            $setter = $extensions.GetMethod('SetAccessControl', [type[]]@($itemObject.GetType(), $securityObject.GetType()))
            if ($null -eq $setter) {
                throw [System.InvalidOperationException]::new('The current Windows runtime cannot apply this filesystem access descriptor.')
            }
            $null = $setter.Invoke($null, [object[]]@($itemObject, $securityObject))
        }
        return
    }
    $getter = [System.IO.File].GetMethod('GetUnixFileMode', [type[]]@([string]))
    $setter = [System.IO.File].GetMethods() | Where-Object {
        $_.Name -eq 'SetUnixFileMode' -and $_.GetParameters()[0].ParameterType -eq [string]
    } | Select-Object -First 1
    if ($null -eq $getter -or $null -eq $setter) {
        throw [System.PlatformNotSupportedException]::new('Unix staging protection requires PowerShell 7.2/.NET 6 or newer.')
    }
    if ([string]::IsNullOrEmpty($Source)) {
        $mode = [System.Enum]::ToObject($getter.ReturnType, 448)
    } else {
        $mode = $getter.Invoke($null, [object[]]@($Source))
    }
    $null = $setter.Invoke($null, [object[]]@($Path, $mode))
}

function New-OutputTransaction {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private staging adapter called only after the archive or restore operation passes ShouldProcess.') ]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $Destination)
    $parent = [System.IO.Path]::GetDirectoryName($Destination)
    $null = [System.IO.Directory]::CreateDirectory($parent)
    $prefix = [System.IO.Path]::Combine($parent, '.' + [System.IO.Path]::GetFileName($Destination) + '.embelish')
    Assert-EmbelishPath -Path $parent
    $lockPath = $prefix + '.lock'
    if (Test-Path -LiteralPath $lockPath) {
        $lockItem = Get-Item -LiteralPath $lockPath -Force
        if ($lockItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint -or $lockItem.LinkType -eq 'HardLink') {
            throw [System.IO.IOException]::new('The publication lock must be an ordinary owned file.')
        }
    }
    $lock = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::OpenOrCreate,
        [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $transaction = $prefix + '.transaction'
    $created = $false
    try {
        if (Test-Path -LiteralPath $transaction) {
            throw [System.IO.IOException]::new('An interrupted output transaction needs recovery; inspect the sibling .embelish.transaction directory before retrying.')
        }
        $null = [System.IO.Directory]::CreateDirectory($transaction)
        $created = $true
        Set-OutputAccess -Path $transaction
        foreach ($name in @('native', 'incoming', 'ready')) {
            $child = [System.IO.Path]::Combine($transaction, $name)
            $null = [System.IO.Directory]::CreateDirectory($child)
            Set-OutputAccess -Path $child
        }
        return [pscustomobject]@{
            Destination = $Destination; Path = $transaction; Lock = $lock
            Native = [System.IO.Path]::Combine($transaction, 'native')
            Incoming = [System.IO.Path]::Combine($transaction, 'incoming')
            Ready = [System.IO.Path]::Combine($transaction, 'ready')
            Previous = [System.IO.Path]::Combine($transaction, 'previous')
            Publishing = $false; BackupPath = $null
        }
    } catch {
        if ($created) {
            try {
                [System.IO.Directory]::Delete($transaction, $true)
            } catch {
                Write-Warning 'Owned staging initialization cleanup failed; inspect the transaction directory before retrying.'
            }
        }
        $lock.Dispose()
        throw
    }
}

function Copy-OutputTree {
    [CmdletBinding()]
    param([string] $Source, [string] $Destination, [object[]] $Inventory, [switch] $PreserveAccess)
    foreach ($entry in $Inventory) {
        $target = [System.IO.Path]::Combine($Destination, $entry.Name)
        if ($entry.Kind -eq 'Directory') {
            $null = [System.IO.Directory]::CreateDirectory($target)
        } else {
            $inputStream = [System.IO.File]::OpenRead([System.IO.Path]::Combine($Source, $entry.Name))
            $outputStream = $null
            try {
                $outputStream = [System.IO.File]::Open($target, [System.IO.FileMode]::Create,
                    [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                $inputStream.CopyTo($outputStream, 81920)
            } finally {
                if ($null -ne $outputStream) { $outputStream.Dispose() }
                $inputStream.Dispose()
            }
        }
        if ($PreserveAccess) {
            Set-OutputAccess -Path $target -Source ([System.IO.Path]::Combine($Source, $entry.Name))
        }
    }
}

function Publish-OutputTransaction {
    [CmdletBinding()]
    param([object] $Transaction, [bool] $AllowExisting, [string] $ReplaceBase)
    $old = @()
    $exists = [System.IO.Directory]::Exists($Transaction.Destination)
    Assert-EmbelishPath -Path $Transaction.Destination
    if ($exists) {
        $old = @(Get-OutputInventory -Path $Transaction.Destination)
        if ($old.Count -gt 0 -and -not $AllowExisting) {
            throw [System.IO.IOException]::new('The destination changed or became nonempty; output was not replaced.')
        }
        Set-OutputAccess -Path $Transaction.Ready -Source $Transaction.Destination
        Copy-OutputTree -Source $Transaction.Destination -Destination $Transaction.Ready -Inventory $old -PreserveAccess
        $copied = @(Get-OutputInventory -Path $Transaction.Ready)
        if (($old | ConvertTo-Json -Depth 4 -Compress) -cne ($copied | ConvertTo-Json -Depth 4 -Compress)) {
            throw [System.IO.IOException]::new('The destination changed while preparing publication; original output was retained.')
        }
    }
    if (-not [string]::IsNullOrEmpty($ReplaceBase)) {
        $pattern = '^' + [regex]::Escape($ReplaceBase) + '\.[0-9]+\.txt$'
        $matching = [System.Text.RegularExpressions.RegexOptions]::None
        if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
            $matching = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        }
        foreach ($file in @(Get-ChildItem -LiteralPath $Transaction.Ready -File -Force)) {
            if ([regex]::IsMatch($file.Name, $pattern, $matching)) {
                [System.IO.File]::Delete($file.FullName)
            }
        }
    }
    $incoming = @(Get-OutputInventory -Path $Transaction.Incoming)
    Copy-OutputTree -Source $Transaction.Incoming -Destination $Transaction.Ready -Inventory $incoming
    foreach ($entry in $incoming) {
        $previousPath = [System.IO.Path]::Combine($Transaction.Destination, $entry.Name)
        if ($exists -and (Test-Path -LiteralPath $previousPath -PathType Leaf) -and $entry.Kind -eq 'File') {
            Set-OutputAccess -Path ([System.IO.Path]::Combine($Transaction.Ready, $entry.Name)) -Source $previousPath
        }
    }
    [System.IO.Directory]::Delete($Transaction.Native, $true)
    [System.IO.Directory]::Delete($Transaction.Incoming, $true)
    $Transaction.Publishing = $true
    if ($exists) {
        [System.IO.Directory]::Move($Transaction.Destination, $Transaction.Previous)
        $moved = @(Get-OutputInventory -Path $Transaction.Previous)
        if (($old | ConvertTo-Json -Depth 4 -Compress) -cne ($moved | ConvertTo-Json -Depth 4 -Compress)) {
            throw [System.IO.IOException]::new('The destination changed during publication; previous output is retained for recovery.')
        }
    }
    [System.IO.Directory]::Move($Transaction.Ready, $Transaction.Destination)
    if ($exists) {
        $Transaction.BackupPath = $Transaction.Path + '.backup-' + [guid]::NewGuid().ToString('N')
        [System.IO.Directory]::Move($Transaction.Path, $Transaction.BackupPath)
    } else {
        [System.IO.Directory]::Delete($Transaction.Path, $true)
    }
    $Transaction.Publishing = $false
}

function Close-OutputTransaction {
    [CmdletBinding()]
    param([object] $Transaction)
    try {
        if ($Transaction.Publishing) {
            if (-not [System.IO.Directory]::Exists($Transaction.Destination) -and
                [System.IO.Directory]::Exists($Transaction.Previous)) {
                [System.IO.Directory]::Move($Transaction.Previous, $Transaction.Destination)
            }
            Write-Warning 'Publication did not finish cleanly. Retained transaction data needs inspection before retrying.'
        } elseif ([System.IO.Directory]::Exists($Transaction.Path)) {
            [System.IO.Directory]::Delete($Transaction.Path, $true)
        }
    } catch {
        Write-Warning 'Owned staging cleanup or rollback could not finish; retained output needs inspection.'
    } finally {
        $Transaction.Lock.Dispose()
    }
}

function Invoke-EmbelishOperation {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string] $SevenZip, [hashtable] $Options)
    if (-not (Test-Path -LiteralPath $Options.SourcePath -PathType Container)) {
        throw [System.ArgumentException]::new('Embelish requires an existing source directory.')
    }
    $source = (Get-Item -LiteralPath $Options.SourcePath -Force).FullName
    $destination = Get-EmbelishDestination -Source $source -Destination $Options.DestinationPath -AllowExisting $Options.Force
    $fields = [ordered]@{
        Operation = 'Embelish'; Succeeded = $false; WhatIf = $true; Status = 'WhatIf'
        Source = $source; Destination = $destination; SevenZipPath = $SevenZip
        SevenZipExit = $null; ShardPattern = $Options.ArchiveBaseName + '.7z.*.txt'
    }
    if (-not $PSCmdlet.ShouldProcess($destination, 'Create and publish encrypted base64 shards')) {
        $script:OperationResult = ConvertTo-ResultRecord -Fields $fields
        return
    }
    $transaction = New-OutputTransaction -Destination $destination
    try {
        $archive = [System.IO.Path]::Combine($transaction.Native, $Options.ArchiveBaseName + '.7z')
        $arguments = @('a', '-t7z', '-mhe=on', ('-mx={0}' -f $Options.CompressionLevel),
            ('-v{0}m' -f $Options.VolumeSizeMegabytes), '-p', '-y', '--', $archive, $source)
        Write-DiagnosticEvent -Stage 'embelish.start' -Data @{ volumeMb = $Options.VolumeSizeMegabytes; mx = $Options.CompressionLevel }
        Write-Verbose 'Running 7-Zip; enter the archive password when prompted.'
        Invoke-SevenZipCommand -Executable $SevenZip -ArgumentList $arguments
        $outcome = Resolve-SevenZipOutcome -ExitCode $script:LastSevenZipExit
        Write-DiagnosticEvent -Stage 'embelish.encrypted' -Data @{ exit = $outcome.ExitCode }
        Assert-SevenZipOutcome -Outcome $outcome
        $volumes = @(Get-OrderedVolume -Files @(Get-ChildItem -LiteralPath $transaction.Native -File -Force) -BaseName ($Options.ArchiveBaseName + '.7z'))
        foreach ($volume in $volumes) {
            $target = [System.IO.Path]::Combine($transaction.Incoming, $volume.Name + '.txt')
            ConvertTo-Base64Shard -VolumePath $volume.FullName -ShardPath $target
            if ((Get-Item -LiteralPath $target -Force).Length -ne (4 * [math]::Ceiling($volume.Length / 3.0))) {
                throw [System.IO.IOException]::new('A staged shard failed its length postcondition.')
            }
        }
        Publish-OutputTransaction -Transaction $transaction -AllowExisting $Options.Force -ReplaceBase ($Options.ArchiveBaseName + '.7z')
        $fields.Succeeded = $true
        $fields.WhatIf = $false
        $fields.Status = $outcome.Status
        $fields.SevenZipExit = $outcome.ExitCode
        $fields['ShardCount'] = $volumes.Count
        $fields['BackupPath'] = $transaction.BackupPath
    } finally {
        Close-OutputTransaction -Transaction $transaction
    }
    Write-DiagnosticEvent -Stage 'embelish.done' -Data @{ shards = $volumes.Count }
    $script:OperationResult = ConvertTo-ResultRecord -Fields $fields
}

function Invoke-DeembelishOperation {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string] $SevenZip, [hashtable] $Options)
    $manifest = Resolve-ShardManifest -Path $Options.SourcePath
    $destination = Get-EmbelishDestination -Source $manifest.Directory -Destination $Options.DestinationPath -AllowExisting $Options.Force
    $fields = [ordered]@{
        Operation = 'Deembelish'; Succeeded = $false; WhatIf = $true; Status = 'WhatIf'
        Source = $manifest.Directory; Destination = $destination; SevenZipPath = $SevenZip
        SevenZipExit = $null; ShardCount = $manifest.Shards.Count
    }
    if (-not $PSCmdlet.ShouldProcess($destination, 'Extract into staging and publish restored files')) {
        $script:OperationResult = ConvertTo-ResultRecord -Fields $fields
        return
    }
    $transaction = New-OutputTransaction -Destination $destination
    try {
        foreach ($shard in $manifest.Shards) {
            $target = [System.IO.Path]::Combine($transaction.Native, [System.IO.Path]::GetFileNameWithoutExtension($shard.Name))
            ConvertFrom-Base64Shard -ShardPath $shard.FullName -VolumePath $target
        }
        $first = [System.IO.Path]::Combine($transaction.Native, $manifest.BaseName + '.001')
        $arguments = @('x', '-aoa', ('-o{0}' -f $transaction.Incoming), '--', $first)
        Write-DiagnosticEvent -Stage 'deembelish.start' -Data @{ shards = $manifest.Shards.Count }
        Write-Verbose 'Running 7-Zip; enter the archive password when prompted.'
        Invoke-SevenZipCommand -Executable $SevenZip -ArgumentList $arguments
        $outcome = Resolve-SevenZipOutcome -ExitCode $script:LastSevenZipExit
        Write-DiagnosticEvent -Stage 'deembelish.extracted' -Data @{ exit = $outcome.ExitCode }
        Assert-SevenZipOutcome -Outcome $outcome
        $restored = @(Get-OutputInventory -Path $transaction.Incoming)
        if ($restored.Count -eq 0) {
            throw [System.IO.IOException]::new('7-Zip returned no restored entries; output was not published.')
        }
        Publish-OutputTransaction -Transaction $transaction -AllowExisting $Options.Force
        $fields.Succeeded = $true
        $fields.WhatIf = $false
        $fields.Status = $outcome.Status
        $fields.SevenZipExit = $outcome.ExitCode
        $fields['RestoredFileCount'] = @($restored | Where-Object { $_.Kind -eq 'File' }).Count
        $fields['BackupPath'] = $transaction.BackupPath
    } finally {
        Close-OutputTransaction -Transaction $transaction
    }
    Write-DiagnosticEvent -Stage 'deembelish.done' -Data @{ restored = $fields.RestoredFileCount }
    $script:OperationResult = ConvertTo-ResultRecord -Fields $fields
}

function ConvertTo-ResultRecord {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Specialized.OrderedDictionary] $Fields
    )

    $Fields['RunId'] = $script:RunId
    $Fields['LogPath'] = $script:LogPath
    return [pscustomobject]$Fields
}

function Resolve-ExitCodeFromError {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord] $ErrorRecord
    )

    $exception = $ErrorRecord.Exception.GetBaseException()
    if ($exception -is [System.ArgumentException] -or $exception -is [System.FormatException]) {
        return 2
    }
    if ($exception -is [System.IO.IOException] -or
        $exception -is [System.UnauthorizedAccessException] -or
        $exception -is [System.Security.SecurityException] -or
        $exception -is [System.InvalidOperationException] -or
        $exception -is [System.ComponentModel.Win32Exception]) {
        return 3
    }
    return 1
}

$script:DebugEnabled = [bool]$PSBoundParameters['Debug']
if ($script:DebugEnabled) {

    $DebugPreference = 'Continue'
}

Initialize-DiagnosticContext -RequestedDirectory $LogDirectory
Write-DiagnosticEvent -Stage 'run.start' -Data @{
    mode = $Mode; psVersion = $PSVersionTable.PSVersion.ToString(); edition = $PSVersionTable.PSEdition
}

try {
    $sevenZip = Resolve-SevenZipPath -ExplicitPath $SevenZipPath

    $options = @{
        SourcePath          = $SourcePath
        DestinationPath     = $DestinationPath
        VolumeSizeMegabytes = $VolumeSizeMegabytes
        CompressionLevel    = $CompressionLevel
        ArchiveBaseName     = $ArchiveBaseName
        Force               = [bool]$Force
    }

    if ($Mode -eq 'Embelish') {
        Invoke-EmbelishOperation -SevenZip $sevenZip -Options $options
    } else {
        Invoke-DeembelishOperation -SevenZip $sevenZip -Options $options
    }

    $result = $script:OperationResult
    Write-Output $result
    Write-DiagnosticEvent -Stage 'run.end' -Data @{ succeeded = $result.Succeeded; whatIf = $result.WhatIf }
    exit 0
} catch {
    $code = Resolve-ExitCodeFromError -ErrorRecord $_
    $kind = switch ($code) { 2 { 'Input' } 3 { 'Operational' } default { 'Unexpected' } }
    $hresult = '0x{0:X8}' -f $_.Exception.GetBaseException().HResult
    Write-DiagnosticEvent -Stage 'run.error' -Level 'error' -Data @{ exitCode = $code; errorKind = $kind; line = $_.InvocationInfo.ScriptLineNumber; hresult = $hresult }

    $ErrorActionPreference = 'Continue'
    Write-Error -Message ('Embelish failed (exit ' + $code + ', HRESULT ' + $hresult + '). Previous output and retained recovery data must be inspected before retrying. Native 7-Zip diagnostics, if any, appear above.') -ErrorId ('Embelish' + $kind) -Category OperationStopped
    exit $code
}
