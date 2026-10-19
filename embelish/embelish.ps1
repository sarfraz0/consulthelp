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
$script:ToolVersion = '1.1.1'
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
        Write-Debug ("[{0}] {1}" -f $Stage, $Level)
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
        stage = $Stage
    }
    if ($PSBoundParameters.ContainsKey('Data') -and $null -ne $Data) {
        foreach ($key in $Data.Keys) {
            $record[[string]$key] = $Data[$key]
        }
    }

    try {
        $line = ($record | ConvertTo-Json -Depth 6 -Compress) + "`n"
        [System.IO.File]::AppendAllText($script:LogPath, $line, [System.Text.Encoding]::UTF8)
    } catch {
        Write-Warning ("embelish: unable to write diagnostic event '{0}': {1}" -f $Stage, $_.Exception.Message)
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
        Write-Warning ("embelish: diagnostics disabled, cannot prepare log directory '{0}': {1}" -f $directory, $_.Exception.Message)
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

    $ErrorActionPreference = 'Continue'
    if (Test-Path -Path 'variable:PSNativeCommandUseErrorActionPreference') {
        Set-Variable -Name 'PSNativeCommandUseErrorActionPreference' -Value $false -Scope Local
    }

    & $Executable @ArgumentList
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
        IsSuccess = ($category -eq 'ok' -or $category -eq 'warning')
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
    if ($Outcome.Category -eq 'operational') {
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

    $bytes = [System.IO.File]::ReadAllBytes($VolumePath)
    $text = [System.Convert]::ToBase64String($bytes)
    [System.IO.File]::WriteAllText($ShardPath, $text, [System.Text.Encoding]::ASCII)
}

function ConvertFrom-Base64Shard {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $ShardPath,

        [Parameter(Mandatory = $true)]
        [string] $VolumePath
    )

    $text = [System.IO.File]::ReadAllText($ShardPath, [System.Text.Encoding]::ASCII)
    $bytes = [System.Convert]::FromBase64String($text.Trim())
    [System.IO.File]::WriteAllBytes($VolumePath, $bytes)
}

function Resolve-ShardManifest {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        $directory = [System.IO.Path]::GetDirectoryName((Resolve-Path -LiteralPath $Path).Path)
    } elseif (Test-Path -LiteralPath $Path -PathType Container) {
        $directory = (Resolve-Path -LiteralPath $Path).Path
    } else {
        throw [System.ArgumentException]::new("Source '$Path' does not exist.")
    }

    $pattern = '^(?<base>.+\.7z)\.(?<num>\d+)\.txt$'
    $shards = @(Get-ChildItem -LiteralPath $directory -Filter '*.7z.*.txt' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $pattern })
    if ($shards.Count -eq 0) {
        throw [System.ArgumentException]::new("No '<base>.7z.NNN.txt' base64 shards found in '$directory'.")
    }

    $bases = @($shards | ForEach-Object { [regex]::Match($_.Name, $pattern).Groups['base'].Value } | Sort-Object -Unique)
    if ($bases.Count -gt 1) {
        throw [System.ArgumentException]::new("Multiple shard sets found in '$directory' ($($bases -join ', ')); point -SourcePath at one '.001.txt' shard.")
    }
    $baseName = $bases[0]

    $ordered = @($shards |
            Where-Object { [regex]::Match($_.Name, $pattern).Groups['base'].Value -eq $baseName } |
            Sort-Object -Property @{ Expression = { [int][regex]::Match($_.Name, $pattern).Groups['num'].Value } })
    $firstNumber = [int][regex]::Match($ordered[0].Name, $pattern).Groups['num'].Value
    if ($firstNumber -ne 1) {
        throw [System.ArgumentException]::new("First shard for '$baseName' is volume $firstNumber, not 001; the shard set is incomplete.")
    }

    return [pscustomobject]@{
        Directory = $directory
        BaseName  = $baseName
        Shards    = $ordered
    }
}

function Invoke-EmbelishOperation {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string] $SevenZip,

        [Parameter(Mandatory = $true)]
        [hashtable] $Options
    )

    $source = $Options.SourcePath
    if (-not (Test-Path -LiteralPath $source -PathType Container)) {
        throw [System.ArgumentException]::new("Embelish source '$source' is not an existing folder.")
    }
    $resolvedSource = (Resolve-Path -LiteralPath $source).Path

    $destination = New-CleanOutputDirectory -Path $Options.DestinationPath -AllowExisting:$Options.Force
    $shardGlob = $Options.ArchiveBaseName + '.7z.*.txt'

    $existingShards = @(Get-ChildItem -LiteralPath $destination -Filter $shardGlob -File -ErrorAction SilentlyContinue)
    if ($existingShards.Count -gt 0) {
        if (-not $Options.Force) {
            throw [System.ArgumentException]::new("Base64 shards for '$($Options.ArchiveBaseName)' already exist in '$destination'; pass -Force to overwrite.")
        }
        foreach ($shard in $existingShards) {
            if ($PSCmdlet.ShouldProcess($shard.FullName, 'Remove existing shard')) {
                Remove-Item -LiteralPath $shard.FullName -Force
            }
        }
    }

    Write-DiagnosticEvent -Stage 'embelish.start' -Data @{
        volumeMb = $Options.VolumeSizeMegabytes; mx = $Options.CompressionLevel
        sourceLeaf = [System.IO.Path]::GetFileName($resolvedSource.TrimEnd([char]'\', [char]'/'))
    }

    $target = "folder '$resolvedSource' into base64 shards in '$destination'"
    if (-not $PSCmdlet.ShouldProcess($target, 'Embelish')) {
        $script:OperationResult = ConvertTo-ResultRecord -Fields ([ordered]@{
                Operation    = 'Embelish'
                Succeeded    = $false
                WhatIf       = $true
                Status       = 'WhatIf'
                Source       = $resolvedSource
                Destination  = $destination
                SevenZipPath = $SevenZip
                SevenZipExit = $null
                ShardPattern = $shardGlob
            })
        return
    }

    $staging = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("{0}-{1}-embelish" -f $script:ToolName, $script:RunId))
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    try {
        $archivePath = [System.IO.Path]::Combine($staging, ($Options.ArchiveBaseName + '.7z'))
        $arguments = @(
            'a', '-t7z', '-mhe=on',
            ("-mx={0}" -f $Options.CompressionLevel),
            ("-v{0}m" -f $Options.VolumeSizeMegabytes),
            '-p', '-y', '--', $archivePath, $resolvedSource
        )

        Write-Verbose "Running 7-Zip; enter the archive password when prompted (twice)."
        Invoke-SevenZipCommand -Executable $SevenZip -ArgumentList $arguments
        $exitCode = $script:LastSevenZipExit
        $outcome = Resolve-SevenZipOutcome -ExitCode $exitCode
        Write-DiagnosticEvent -Stage 'embelish.encrypted' -Level ($(if ($outcome.IsSuccess) { 'info' } else { 'error' })) -Data @{ exit = $exitCode; status = $outcome.Status }
        Assert-SevenZipOutcome -Outcome $outcome
        if ($outcome.IsWarning) {
            Write-Warning "7-Zip reported a non-fatal warning (exit 1); the archive was created but review the output above."
        }

        $volumes = @(Get-ChildItem -LiteralPath $staging -Filter ($Options.ArchiveBaseName + '.7z.*') -File -ErrorAction SilentlyContinue |
                Sort-Object -Property Name)
        foreach ($volume in $volumes) {
            $shardPath = [System.IO.Path]::Combine($destination, ($volume.Name + '.txt'))
            ConvertTo-Base64Shard -VolumePath $volume.FullName -ShardPath $shardPath

            Remove-Item -LiteralPath $volume.FullName -Force
        }
        $shardCount = $volumes.Count
    } finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-DiagnosticEvent -Stage 'embelish.done' -Data @{ shards = $shardCount }
    $script:OperationResult = ConvertTo-ResultRecord -Fields ([ordered]@{
            Operation    = 'Embelish'
            Succeeded    = $true
            WhatIf       = $false
            Status       = $outcome.Status
            Source       = $resolvedSource
            Destination  = $destination
            SevenZipPath = $SevenZip
            SevenZipExit = $exitCode
            ShardPattern = $shardGlob
            ShardCount   = $shardCount
        })
}

function Invoke-DeembelishOperation {
    # .EXTERNALHELP ../docs/embelish/embelish-help.xml
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string] $SevenZip,

        [Parameter(Mandatory = $true)]
        [hashtable] $Options
    )

    $manifest = Resolve-ShardManifest -Path $Options.SourcePath
    $destination = New-CleanOutputDirectory -Path $Options.DestinationPath -AllowExisting:$Options.Force
    $shardCount = $manifest.Shards.Count

    Write-DiagnosticEvent -Stage 'deembelish.start' -Data @{
        base = $manifest.BaseName; shards = $shardCount
    }

    $target = "$shardCount base64 shard(s) for '$($manifest.BaseName)' into folder '$destination'"
    if (-not $PSCmdlet.ShouldProcess($target, 'Deembelish')) {
        $script:OperationResult = ConvertTo-ResultRecord -Fields ([ordered]@{
                Operation    = 'Deembelish'
                Succeeded    = $false
                WhatIf       = $true
                Status       = 'WhatIf'
                Source       = $manifest.Directory
                Destination  = $destination
                SevenZipPath = $SevenZip
                SevenZipExit = $null
                ShardCount   = $shardCount
            })
        return
    }

    $staging = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("{0}-{1}-deembelish" -f $script:ToolName, $script:RunId))
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    try {
        foreach ($shard in $manifest.Shards) {
            $volumeName = [System.IO.Path]::GetFileNameWithoutExtension($shard.Name)
            $volumePath = [System.IO.Path]::Combine($staging, $volumeName)
            ConvertFrom-Base64Shard -ShardPath $shard.FullName -VolumePath $volumePath
        }
        $firstVolume = [System.IO.Path]::Combine($staging, ($manifest.BaseName + '.001'))

        $arguments = @('x', '-aoa', ("-o{0}" -f $destination), '--', $firstVolume)

        Write-Verbose "Running 7-Zip; enter the archive password when prompted."
        Invoke-SevenZipCommand -Executable $SevenZip -ArgumentList $arguments
        $exitCode = $script:LastSevenZipExit
        $outcome = Resolve-SevenZipOutcome -ExitCode $exitCode
        Write-DiagnosticEvent -Stage 'deembelish.extracted' -Level ($(if ($outcome.IsSuccess) { 'info' } else { 'error' })) -Data @{ exit = $exitCode; status = $outcome.Status }
        Assert-SevenZipOutcome -Outcome $outcome
        if ($outcome.IsWarning) {
            Write-Warning "7-Zip reported a non-fatal warning (exit 1); extraction completed but review the output above."
        }
    } finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }

    $restoredCount = @(Get-ChildItem -LiteralPath $destination -Recurse -File -ErrorAction SilentlyContinue).Count
    Write-DiagnosticEvent -Stage 'deembelish.done' -Data @{ restored = $restoredCount }
    $script:OperationResult = ConvertTo-ResultRecord -Fields ([ordered]@{
            Operation         = 'Deembelish'
            Succeeded         = $true
            WhatIf            = $false
            Status            = $outcome.Status
            Source            = $manifest.Directory
            Destination       = $destination
            SevenZipPath      = $SevenZip
            SevenZipExit      = $exitCode
            ShardCount        = $shardCount
            RestoredFileCount = $restoredCount
        })
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

    $exception = $ErrorRecord.Exception
    if ($exception -is [System.ArgumentException]) {
        return 2
    }
    if ($exception -is [System.IO.IOException] -or
        $exception -is [System.UnauthorizedAccessException] -or
        $exception -is [System.Security.SecurityException] -or
        $exception -is [System.InvalidOperationException]) {
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
    Write-DiagnosticEvent -Stage 'run.error' -Level 'error' -Data @{ exitCode = $code; message = $_.Exception.Message }

    $ErrorActionPreference = 'Continue'
    Write-Error $_
    exit $code
}
