#requires -Version 5.1
<#
.SYNOPSIS
    Plans and resumes a paced transfer of archive shards to an approved Git remote.
.DESCRIPTION
    Revision 1.0.0. One self-contained Windows PowerShell 5.1 script. Requires
    installed Git for Windows 2.30 or newer, FullLanguage, and pre-provisioned credentials.
    Run as the current user. No installation, elevation, policy changes are performed. Git hooks, signing,
    credential helpers, TLS checks, and repository filters remain active.

    Plan reads local repository information and writes a readable UTF-8 JSON
    document. It performs no network operation, staging, commit, or push. The
    document records SHA-256 hashes, exact Git blob IDs, the branch, the configured
    push destination's fingerprint, and a fixed wait for each pending file.
    Pacing is transparent and deterministic.

    Push revalidates that document and processes one pending file at a time:
    verify -> add -> verify index -> commit -> verify commit -> push -> verify
    remote branch -> checkpoint. All shard sets must have contiguous numbering.
    Each matching checksum manifest is included after that set's shards. Legacy
    sets are accepted with a warning. This does not decrypt or test the archive.
    Already-identical files in the base commit need no new commit.

    Push starts a 60-minute budget, persisted at <PlanPath>.state.json. The same
    Push command resumes automatically. Verified Git commits and the actual
    remote branch, rather than an optimistic counter, establish progress. One
    local commit may be awaiting push. Failed/uncertain pushes are not blindly
    retried. Last attempted/verified push times preserve pacing on resume.
    A timed-out plan remains inspectable but cannot restart its original budget.
    A completed checkpoint may be reverified later without sending more data.
    A new Plan is an explicit new transfer, with a new budget.
    Its base still must match the remote: resolve any pending older local commit
    through the approved Git recovery procedure before starting a replacement plan.

    Completion within 60 minutes cannot be guaranteed: counts, bandwidth,
    authentication, hooks, server limits, and outages can make it impossible.
    Plan rejects waits consuming more than 50 minutes, reserving at least ten
    minutes for work. Push bounds Git calls and stops initiating work at its
    deadline. Blocking filesystem I/O and process cleanup can exceed the budget.
    Keep the Windows clock accurate. Fixed waiting starts after the previous
    attempt/verified completion; operation time adds to the planned waits.

    The current branch must already exist remotely. At the first Push its remote
    tip must equal the plan's base commit, preventing upload of unrelated local
    commits. Divergence, branch changes, changed source bytes, and foreign staged
    changes cause a stop. Do not edit or use Git in this worktree during transfer.
    The exclusive sendshards.lock in Git's common directory serializes this tool
    across linked worktrees; it cannot lock out other Git applications.

    Plan/state/logs must be outside the worktree and Git metadata. Plans contain
    private paths and file hashes and must be reviewed before sharing. Logs omit
    paths, remote URLs, payloads, credentials, raw Git output, and exception text.
    No success-stream result is produced. Files/shards are never removed. Failed
    commits can leave the next file staged; successful commits remain in Git.
    A verified, matching staged next file is accepted on resume. Rewriting history,
    resetting, stashing, force-pushing, and silently skipping hooks are not used.
    Hard termination can also leave a native Git index.lock; investigate native
    Git locks before resuming rather than deleting them while Git may still run.
    Use actual directory paths for plans/logs: directory symlink/junction ancestry
    is refused for these owned outputs so an external-looking path cannot redirect
    writes into the repository. Ordinary OneDrive cloud placeholders are not banned.
    Filters changing shard bytes (including LFS pointers) fail verification;
    use a repository configuration approved for preserving these file bytes.

    Runtime limits: 10,000 shard/manifest files, 256 MiB per shard, 8 MiB per JSON
    document, 2 MiB captured output per native stream, 100 log files retained.
    Child executables: installed git.exe and its configured approved helpers/hooks;
    Windows taskkill.exe only for the process tree owned by this invocation on
    timeout/cancellation. No network contact occurs until an authorized Push.
.PARAMETER Plan
    Generate a new plan. Existing plan or checkpoint files are never overwritten.
.PARAMETER Push
    Execute or resume an existing plan. Repeating this command is the resume action.
.PARAMETER SourcePath
    With Plan: a local folder inside the worktree, containing *.7z.NNN.txt shards.
    Top-level files only; multiple archive sets are supported. Other files ignored.
.PARAMETER PlanPath
    Required in either mode. JSON document outside the repository. Relative paths
    use the invocation's PowerShell filesystem location, not the script directory.
.PARAMETER Remote
    With Plan: configured remote name, default origin. Exactly one push URL required.
    The configured push URL is used for both push and remote verification.
.PARAMETER WaitSeconds
    With Plan: fixed minimum wait between pushes, default 30, range 5 through 600.
    This is a rate limit, not a means of avoiding security inspection.
.PARAMETER GitTimeoutSeconds
    With Plan: maximum duration of each Git command, default 60, range 10 through 300,
    further limited by the remaining execution budget.
.PARAMETER GitPath
    Optional approved installed git.exe path. Otherwise resolved from PATH.
.EXAMPLE
    .\sendshards.ps1 -Plan -SourcePath 'C:\repo\shards' -PlanPath 'C:\TransferPlans\batch.json'
.EXAMPLE
    .\sendshards.ps1 -Push -PlanPath 'C:\TransferPlans\batch.json'
    Repeat the same command after an interruption; no separate Resume flag is needed.
.EXAMPLE
    .\sendshards.ps1 -Push -PlanPath 'C:\TransferPlans\batch.json' -WhatIf
.EXAMPLE
    .\sendshards.ps1 -Push -PlanPath 'C:\TransferPlans\batch.json' -Debug
.INPUTS
    None. Source files and the persisted plan are input data.
.OUTPUTS
    None on the success stream. Plan produces a JSON document. Push maintains a
    checkpoint and sanitized JSONL operational logs. Verbose/Debug use their streams.
.NOTES
    WhatIf/declined confirmation: read-only local checks, no network, logs, locks,
    directories, staging, or checkpoint writes. Debug never introduces prompting.
    Exit: 0 complete/requested preview; 1 unexpected; 2 invalid input/plan;
    3 operational/dependency/repository conflict; 4 controlled timeout.
    Binding, execution-policy refusal, Ctrl+C, and hard termination can occur outside
    this mapping. Missing completion is an unknown outcome; resume reconciles Git.
    Audit initialization must succeed before mutations. Later logging/cleanup failure
    warns without replacing the primary failure. Checkpoint failure stops further work.
    Normal retention removes only this tool's old log files, never plans/checkpoints.

    Task Scheduler example (paths illustrative):
      Program: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
      Arguments: -NoLogo -NoProfile -NonInteractive -File "C:\ApprovedTools\sendshards.ps1" -Push -PlanPath "C:\TransferPlans\batch.json"
      Start in: C:\ApprovedTools
    Use the intended user with approved network credentials; S4U lacks network access.
    Configure no overlapping instances, no automatic rapid retries, an execution limit
    above 60 minutes for cleanup, and suitable network/sleep/battery conditions. This
    script does not register tasks. External helpers/hooks can require prior setup;
    noninteractive PowerShell does not prevent every external application dialog.

    Verification for revision 1.0.0:
      PASS: development-only native Git protocol checks on Linux, Git 2.51.1,
        Python 3.12: 11 asserted scenarios covering exact file/commit/remote bytes,
        manifest order, distinct fetch/push destinations, interrupted progress,
        Unicode/literal paths, hook rejection, filters, and wait-budget boundaries.
        These exercised native Git commands and a quoting model, NOT this script.
      PASS: lexical source checks (balanced strings/delimiters, unique functions,
        command-resolution inventory, prohibited-operation review), UTF-8 BOM.
      PASS: manual review of function size/control flow/nesting, source/index/commit
        guards, output suppression, and sanitized diagnostics. No AST metric claim.
      NOT RUN: PowerShell parser/PSScriptAnalyzer, target Windows PowerShell 5.1,
        Windows process/ACL behavior, Task Scheduler, endpoint/proxy/DLP acceptance.
        Those runtimes and the managed target environment are unavailable here.
.LINK
    https://git-scm.com/docs/git-push
.LINK
    https://git-scm.com/docs/git-ls-remote
.LINK
    https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/taskkill
#>
[CmdletBinding(DefaultParameterSetName = 'Plan', PositionalBinding = $false,
    SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(ParameterSetName = 'Plan')]
    [switch] $Plan,
    [Parameter(ParameterSetName = 'Push')]
    [switch] $Push,
    [Parameter(ParameterSetName = 'Plan')]
    [string] $SourcePath,
    [string] $PlanPath,
    [Parameter(ParameterSetName = 'Plan')]
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$')]
    [string] $Remote = 'origin',
    [Parameter(ParameterSetName = 'Plan')]
    [ValidateRange(5, 600)]
    [int] $WaitSeconds = 30,
    [Parameter(ParameterSetName = 'Plan')]
    [ValidateRange(10, 300)]
    [int] $GitTimeoutSeconds = 60,
    [string] $GitPath
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'
$debugEnabled = $false
if ($PSBoundParameters.ContainsKey('Debug')) {
    $debugEnabled = [bool]$PSBoundParameters['Debug']
}
$DebugPreference = 'SilentlyContinue'
if ($debugEnabled) {
    $DebugPreference = 'Continue'
}

function Stop-Transfer {
    [CmdletBinding()]
    param([string] $Id, [string] $Message, [int] $Code = 3)
    $exception = [System.InvalidOperationException]::new($Message)
    $exception.Data['TransferId'] = $Id
    $exception.Data['TransferCode'] = $Code
    $record = [System.Management.Automation.ErrorRecord]::new($exception, $Id,
        [System.Management.Automation.ErrorCategory]::InvalidOperation, $null)
    $PSCmdlet.ThrowTerminatingError($record)
}

function Resolve-LocalPath {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        Stop-Transfer -Id 'MissingPath' -Message 'A required filesystem path is missing.' -Code 2
    }
    $provider = $null
    $drive = $null
    $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Value, [ref]$provider, [ref]$drive)
    if ($provider.Name -ne 'FileSystem') {
        Stop-Transfer -Id 'PathProvider' -Message 'Use a filesystem path.' -Code 2
    }
    return [System.IO.Path]::GetFullPath($path)
}

function Test-PathWithin {
    [CmdletBinding()]
    [OutputType([bool])]
    param([string] $Path, [string] $Directory)
    $root = $Directory.TrimEnd([char[]]'\/')
    if ($Path.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    return $Path.StartsWith($root + [System.IO.Path]::DirectorySeparatorChar,
        [System.StringComparison]::OrdinalIgnoreCase)
}

function Assert-OwnedFilePath {
    [CmdletBinding()]
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    $link = $item.PSObject.Properties['LinkType']
    $isLink = $false
    if ($null -ne $link) {
        $isLink = @('SymbolicLink', 'Junction', 'HardLink') -contains $link.Value
    }
    if ($item.PSIsContainer -or $isLink) {
        Stop-Transfer -Id 'FileLink' -Message 'A tool-owned data file must be a regular file.' -Code 2
    }
}

function Get-TextHash {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Text)
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $algorithm.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text))
        return ([System.BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant()
    } finally {
        $algorithm.Dispose()
    }
}

function Get-FileFingerprint {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $Path, [long] $MaximumBytes)
    Assert-OwnedFilePath -Path $Path
    $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read')
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $length = $stream.Length
        if ($length -gt $MaximumBytes) {
            Stop-Transfer -Id 'FileTooLarge' -Message 'An input file exceeds its documented size limit.' -Code 2
        }
        $hash = $algorithm.ComputeHash($stream)
        if ($stream.Length -ne $length) {
            Stop-Transfer -Id 'ChangedFile' -Message 'An input file changed during inspection.' -Code 2
        }
        return [pscustomobject]@{
            Bytes = $length
            Sha256 = ([System.BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant()
        }
    } finally {
        $algorithm.Dispose()
        $stream.Dispose()
    }
}

function ConvertTo-NativeArgument {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string] $Value)
    if ($Value.IndexOf([char]0) -ge 0) {
        Stop-Transfer -Id 'NativeArgument' -Message 'An argument contains an invalid null character.' -Code 2
    }
    # Windows CRT quoting: double slashes before a quote and before the closing quote.
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function Get-RemainingTime {
    [CmdletBinding()]
    [OutputType([double])]
    param([object] $Context)
    if ($null -eq $Context.DeadlineUtc) {
        return [double]::PositiveInfinity
    }
    $wall = ($Context.DeadlineUtc - [datetime]::UtcNow).TotalSeconds
    $monotonic = $Context.AllowedSeconds - $Context.Watch.Elapsed.TotalSeconds
    $remaining = [math]::Min($wall, $monotonic)
    if ($remaining -le 0) {
        Stop-Transfer -Id 'Deadline' -Message 'The transfer budget has expired; progress is retained.' -Code 4
    }
    return $remaining
}

function Receive-NativeStream {
    [CmdletBinding()]
    param([object] $Channel)
    if ($Channel.Complete) {
        return
    }
    if (-not $Channel.ReadTask.IsCompleted) {
        return
    }
    $length = $Channel.ReadTask.GetAwaiter().GetResult()
    if ($length -eq 0) {
        $Channel.Complete = $true
        return
    }
    if ($Channel.Data.Length + $length -gt 2097152) {
        Stop-Transfer -Id 'NativeOutputLimit' -Message 'A native command exceeded the output limit.'
    }
    $Channel.Data.Write($Channel.Buffer, 0, $length)
    $Channel.ReadTask = $Channel.Stream.ReadAsync($Channel.Buffer, 0, $Channel.Buffer.Length)
}

function Stop-OwnedProcess {
    [CmdletBinding()]
    param([System.Diagnostics.Process] $Process, [string] $TaskKillPath)
    if ($Process.HasExited) {
        return
    }
    # Never kill by image name. The PID belongs to the Process object started here.
    $killer = [System.Diagnostics.Process]::new()
    try {
        $killer.StartInfo.FileName = $TaskKillPath
        $killer.StartInfo.Arguments = '/PID ' + $Process.Id + ' /T /F'
        $killer.StartInfo.UseShellExecute = $false
        $killer.StartInfo.CreateNoWindow = $true
        $killer.StartInfo.RedirectStandardOutput = $true
        $killer.StartInfo.RedirectStandardError = $true
        $null = $killer.Start()
        $outTask = $killer.StandardOutput.ReadToEndAsync()
        $errorTask = $killer.StandardError.ReadToEndAsync()
        if (-not $killer.WaitForExit(5000)) {
            $killer.Kill()
            Write-Warning ('Owned-process cleanup exceeded its wait limit. Inspect Git PID: ' + $Process.Id)
        }
        if (-not $Process.WaitForExit(2000)) {
            Write-Warning ('An owned Git process may still be running. PID: ' + $Process.Id)
        }
    } catch {
        Write-Warning ('Owned-process cleanup failed. Inspect Git PID: ' + $Process.Id)
    } finally {
        $killer.Dispose()
    }
}

function Invoke-TransferGit {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([object] $Context, [string[]] $Arguments, [int[]] $AcceptedCodes = @(0))
    $remaining = Get-RemainingTime -Context $Context
    $limit = [math]::Min($Context.GitTimeoutSeconds, $remaining)
    $argumentsList = @('--no-optional-locks', '--literal-pathspecs', '-C', $Context.Root) + $Arguments
    $quoted = @($argumentsList | ForEach-Object {
            ConvertTo-NativeArgument -Value $_
        })
    $process = [System.Diagnostics.Process]::new()
    $channels = @()
    $started = $false
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $process.StartInfo.FileName = $Context.Git
        $process.StartInfo.Arguments = $quoted -join ' '
        $process.StartInfo.WorkingDirectory = $Context.Root
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.CreateNoWindow = $true
        $process.StartInfo.RedirectStandardInput = $true
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true
        $process.StartInfo.EnvironmentVariables['GIT_TERMINAL_PROMPT'] = '0'
        try {
            $started = $process.Start()
        }
        catch {
            Stop-Transfer -Id 'NativeStart' -Message 'The installed Git process could not start; check its path, permissions, and working directory.'
        }
        $process.StandardInput.Close()
        foreach ($stream in @($process.StandardOutput.BaseStream, $process.StandardError.BaseStream)) {
            $buffer = [byte[]]::new(8192)
            $channels += [pscustomobject]@{
                Stream = $stream
                Buffer = $buffer
                Data = [System.IO.MemoryStream]::new()
                Complete = $false
                ReadTask = $stream.ReadAsync($buffer, 0, $buffer.Length)
            }
        }
        Wait-NativeExit -Process $process -Channels $channels -Watch $watch -Limit $limit
        $code = $process.ExitCode
        Write-Debug ('Git stage=' + $Arguments[0] + ' exit=' + $code + ' durationMs=' + $watch.ElapsedMilliseconds)
        Write-TransferEvent -Context $Context -Event 'Git' -Fields @{
            command = $Arguments[0]
            nativeCode = $code
            durationMs = $watch.ElapsedMilliseconds
        } -DebugOnly
        if ($AcceptedCodes -notcontains $code) {
            Stop-Transfer -Id 'GitFailure' -Message ('Git failed during ' + $Arguments[0] + ' (exit ' + $code + '). Raw output is omitted for privacy.')
        }
        $encoding = [System.Text.UTF8Encoding]::new($false, $true)
        return [pscustomobject]@{
            Code = $code
            Text = $encoding.GetString($channels[0].Data.ToArray())
        }
    } finally {
        if ($started) {
            Stop-OwnedProcess -Process $process -TaskKillPath $Context.TaskKill
        }
        foreach ($channel in $channels) {
            $channel.Data.Dispose()
        }
        $process.Dispose()
        $watch.Stop()
    }
}

function Wait-NativeExit {
    [CmdletBinding()]
    param([System.Diagnostics.Process] $Process, [object[]] $Channels,
        [System.Diagnostics.Stopwatch] $Watch, [double] $Limit)
    while (-not ($Process.HasExited -and $Channels[0].Complete -and $Channels[1].Complete)) {
        foreach ($channel in $Channels) {
            Receive-NativeStream -Channel $channel
        }
        if ($Watch.Elapsed.TotalSeconds -ge $Limit) {
            Stop-Transfer -Id 'GitTimeout' -Message 'A Git command timed out; reconcile the retained plan before continuing.' -Code 4
        }
        Start-Sleep -Milliseconds 25
    }
}

function Get-GitText {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Context, [string[]] $Arguments)
    return (Invoke-TransferGit -Context $Context -Arguments $Arguments).Text.TrimEnd([char[]]"`r`n")
}

function Write-TransferEvent {
    [CmdletBinding()]
    param([object] $Context, [string] $Event, [hashtable] $Fields = @{}, [switch] $DebugOnly)
    if ($null -eq $Context.LogPath) {
        return
    }
    if ($DebugOnly -and -not $Context.DebugEnabled) {
        return
    }
    $entry = [ordered]@{
        utc = [datetime]::UtcNow.ToString('o')
        tool = 'sendshards'
        version = '1.0.0'
        runId = $Context.RunId
        event = $Event
    }
    $allowed = @('command', 'nativeCode', 'durationMs', 'fileCount', 'pendingCount', 'verifiedCount',
        'item', 'waitSeconds', 'errorId', 'errorType', 'line', 'status', 'scriptHash',
        'psVersion', 'osVersion', 'architecture', 'languageMode', 'gitVersion', 'planId', 'operation')
    foreach ($key in $Fields.Keys) {
        if ($allowed -contains $key) {
            $entry[$key] = $Fields[$key]
        }
    }
    try {
        $line = ($entry | ConvertTo-Json -Compress -Depth 4) + "`n"
        [System.IO.File]::AppendAllText($Context.LogPath, $line, [System.Text.UTF8Encoding]::new($false))
    } catch {
        Write-Warning 'An operational log write failed; the original outcome is unchanged.'
    }
}

function Initialize-TransferLog {
    [CmdletBinding()]
    param([object] $Context, [string] $Directory)
    $logDirectory = [System.IO.Path]::Combine($Directory, 'sendshards-logs')
    Assert-ControlledDirectory -Path $logDirectory
    $null = [System.IO.Directory]::CreateDirectory($logDirectory)
    $path = [System.IO.Path]::Combine($logDirectory, ('sendshards-' + $Context.RunId + '.jsonl'))
    $stream = [System.IO.File]::Open($path, 'CreateNew', 'Write', 'Read')
    try {
        $entry = [ordered]@{
            utc = [datetime]::UtcNow.ToString('o')
            tool = 'sendshards'
            version = '1.0.0'
            runId = $Context.RunId
            event = 'Start'
            status = 'Running'
        }
        $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes(($entry | ConvertTo-Json -Compress) + "`n")
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    } finally {
        $stream.Dispose()
    }
    $Context.LogPath = $path
    Write-Verbose ('Operational log: ' + $path)
    try {
        $old = @(Get-ChildItem -LiteralPath $logDirectory -File -Filter 'sendshards-*.jsonl' |
            Where-Object {
                $_.Name -match '^sendshards-[0-9a-f]{32}\.jsonl$'
            } | Sort-Object -Property LastWriteTimeUtc -Descending)
        foreach ($file in @($old | Select-Object -Skip 100)) {
            Assert-OwnedFilePath -Path $file.FullName
            Remove-Item -LiteralPath $file.FullName -ErrorAction Stop
        }
    } catch {
        Write-Warning 'Old-log cleanup was incomplete; transfer data was not removed.'
    }
    Write-EnvironmentEvent -Context $Context
}

function Write-EnvironmentEvent {
    [CmdletBinding()]
    param([object] $Context)
    if (-not $Context.DebugEnabled) {
        return
    }
    try {
        $scriptHash = (Get-FileFingerprint -Path $PSCommandPath -MaximumBytes 8388608).Sha256
        Write-TransferEvent -Context $Context -Event 'Environment' -Fields @{
            scriptHash = $scriptHash
            psVersion = $PSVersionTable.PSVersion.ToString()
            osVersion = [System.Environment]::OSVersion.Version.ToString()

            architecture = [System.IntPtr]::Size * 8
            languageMode = $ExecutionContext.SessionState.LanguageMode.ToString()
            gitVersion = $Context.GitVersion
        } -DebugOnly
    } catch {
        Write-Warning 'Some optional environment diagnostics are unavailable.'
    }
}

function Read-TransferJson {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $Path)
    Assert-OwnedFilePath -Path $Path
    $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read')
    $reader = $null
    try {
        if ($stream.Length -gt 8388608) {
            Stop-Transfer -Id 'JsonSize' -Message 'A JSON document exceeds 8 MiB.' -Code 2
        }
        $memory = [System.IO.MemoryStream]::new()
        try {
            $stream.CopyTo($memory)
            $bytes = $memory.ToArray()
        } finally {
            $memory.Dispose()
        }
        $quotas = [System.Xml.XmlDictionaryReaderQuotas]::new()
        $quotas.MaxDepth = 8
        $quotas.MaxStringContentLength = 32768
        $quotas.MaxArrayLength = 10000
        $reader = [System.Runtime.Serialization.Json.JsonReaderWriterFactory]::CreateJsonReader($bytes, $quotas)
        $objects = [System.Collections.Generic.Stack[object]]::new()
        while ($reader.Read()) {
            Test-JsonElement -Reader $reader -Objects $objects
        }
        $text = [System.Text.UTF8Encoding]::new($false, $true).GetString($bytes).TrimStart([char]0xfeff)
        return $text | ConvertFrom-Json -ErrorAction Stop
    } finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        $stream.Dispose()
    }
}

function Test-JsonElement {
    [CmdletBinding()]
    param([System.Xml.XmlDictionaryReader] $Reader, [object] $Objects)
    if ($Reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement) {
        if ($Objects.Count -gt 0 -and $Reader.Depth -eq $Objects.Peek().Depth) {
            $null = $Objects.Pop()
        }
        return
    }
    if ($Reader.NodeType -ne [System.Xml.XmlNodeType]::Element) {
        return
    }
    if ($Objects.Count -gt 0 -and $Reader.Depth -eq $Objects.Peek().Depth + 1) {
        if (-not $Objects.Peek().Names.Add($Reader.LocalName)) {
            Stop-Transfer -Id 'DuplicateJsonKey' -Message 'Duplicate JSON keys are not allowed.' -Code 2
        }
    }
    if ($Reader.GetAttribute('type') -eq 'object' -and -not $Reader.IsEmptyElement) {
        $Objects.Push([pscustomobject]@{
                Depth = $Reader.Depth; Names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            })
    }
}

function Assert-ObjectSchema {
    [CmdletBinding()]
    param([object] $Value, [string[]] $Keys)
    if ($Value -isnot [pscustomobject]) {
        Stop-Transfer -Id 'JsonSchema' -Message 'A JSON object is required.' -Code 2
    }
    $actual = @($Value.PSObject.Properties.Name)
    if ($actual.Count -ne $Keys.Count) {
        Stop-Transfer -Id 'JsonSchema' -Message 'A document has missing or unknown fields.' -Code 2
    }
    foreach ($key in $Keys) {
        if ($actual -cnotcontains $key) {
            Stop-Transfer -Id 'JsonSchema' -Message 'A document has missing or incorrectly cased fields.' -Code 2
        }
    }
}

function Assert-IntegerValue {
    [CmdletBinding()]
    param([object] $Value, [long] $Minimum, [long] $Maximum)
    if ($Value -isnot [int] -and $Value -isnot [long]) {
        Stop-Transfer -Id 'JsonNumber' -Message 'An integer field has an invalid type.' -Code 2
    }
    if ($Value -lt $Minimum -or $Value -gt $Maximum) {
        Stop-Transfer -Id 'JsonNumber' -Message 'An integer field is outside its supported range.' -Code 2
    }
}

function Assert-TextPattern {
    [CmdletBinding()]
    param([object] $Value, [string] $Pattern)
    if ($Value -isnot [string] -or $Value -cnotmatch $Pattern) {
        Stop-Transfer -Id 'JsonString' -Message 'A text field has an invalid type or format.' -Code 2
    }
}

function Write-AtomicJson {
    [CmdletBinding()]
    param([string] $Path, [object] $Document, [switch] $NewOnly)
    Assert-OwnedFilePath -Path $Path
    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes(($Document | ConvertTo-Json -Depth 8) + "`n")
    if ($bytes.Length -gt 8388608) {
        Stop-Transfer -Id 'JsonSize' -Message 'The generated document exceeds 8 MiB.' -Code 2
    }
    $temporary = $Path + '.new-' + [guid]::NewGuid().ToString('N')
    try {
        $stream = [System.IO.FileStream]::new($temporary, [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::Write, [System.IO.FileShare]::None, 4096, [System.IO.FileOptions]::WriteThrough)
        try {
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Flush($true)
        } finally {
            $stream.Dispose()
        }
        $null = Read-TransferJson -Path $temporary
        if ([System.IO.File]::Exists($Path)) {
            if ($NewOnly) {
                Stop-Transfer -Id 'PlanExists' -Message 'The plan already exists; choose a new filename.' -Code 2
            }
            [System.IO.File]::Replace($temporary, $Path, $null)
        } else {
            [System.IO.File]::Move($temporary, $Path)
        }
    } finally {
        if ([System.IO.File]::Exists($temporary)) {
            try {
                [System.IO.File]::Delete($temporary)
            } catch {
                Write-Warning 'An owned temporary JSON file could not be removed.'
            }
        }
    }
}

function New-TransferContext {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $WorkingPath, [string] $ExecutablePath, [int] $Timeout, [bool] $EnableDebug)
    Assert-HostPrerequisite
    if ([string]::IsNullOrWhiteSpace($ExecutablePath)) {
        try {
            $ExecutablePath = (Get-Command -Name 'git.exe' -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
        }
        catch {
            Stop-Transfer -Id 'GitMissing' -Message 'Install approved Git for Windows or provide -GitPath.'
        }
    }
    $executable = Resolve-LocalPath -Value $ExecutablePath
    if (-not [System.IO.File]::Exists($executable)) {
        Stop-Transfer -Id 'GitMissing' -Message 'Supply an approved installed Git executable.'
    }
    $taskKill = [System.IO.Path]::Combine([System.Environment]::SystemDirectory, 'taskkill.exe')
    if (-not [System.IO.File]::Exists($taskKill)) {
        Stop-Transfer -Id 'TaskKillMissing' -Message 'Windows taskkill is unavailable.'
    }
    Add-Type -AssemblyName System.Runtime.Serialization -ErrorAction Stop
    return [pscustomobject]@{
        Root = $WorkingPath
        Git = $executable
        TaskKill = $taskKill
        GitTimeoutSeconds = $Timeout

        DebugEnabled = $EnableDebug
        RunId = [guid]::NewGuid().ToString('N')
        LogPath = $null

        DeadlineUtc = $null
        AllowedSeconds = 3600.0
        Watch = [System.Diagnostics.Stopwatch]::StartNew()

        StartedUtc = [datetime]::UtcNow
        CommonDirectory = $null
        GitDirectory = $null
        PlanHash = $null
        GitVersion = $null
    }
}

function Assert-HostPrerequisite {
    [CmdletBinding()]
    param()
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
        Stop-Transfer -Id 'Platform' -Message 'This script requires Windows: its process-tree cleanup uses taskkill.'
    }
    if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
        Stop-Transfer -Id 'LanguageMode' -Message 'This script requires approved FullLanguage execution.'
    }
    foreach ($name in @('GIT_DIR', 'GIT_WORK_TREE', 'GIT_INDEX_FILE', 'GIT_OBJECT_DIRECTORY', 'GIT_ALTERNATE_OBJECT_DIRECTORIES', 'GIT_NAMESPACE')) {
        if (-not [string]::IsNullOrEmpty([System.Environment]::GetEnvironmentVariable($name))) {
            Stop-Transfer -Id 'GitEnvironment' -Message 'Git repository/index override environment variables are not supported.' -Code 2
        }
    }
}

function Set-RepositoryContext {
    [CmdletBinding()]
    param([object] $Context)
    $Context.GitVersion = Get-GitText -Context $Context -Arguments @('--version')
    Assert-TextPattern -Value $Context.GitVersion -Pattern '^git version [0-9]+\.[0-9]+\.[0-9]+[A-Za-z0-9._-]*$'
    $versionMatch = [regex]::Match($Context.GitVersion, '[0-9]+\.[0-9]+\.[0-9]+')
    if ([version]$versionMatch.Value -lt [version]'2.30.0') {
        Stop-Transfer -Id 'GitVersion' -Message 'Git for Windows 2.30 or newer is required.'
    }
    $root = Get-GitText -Context $Context -Arguments @('rev-parse', '--show-toplevel')
    $Context.Root = Resolve-LocalPath -Value $root
    $gitDirectory = Get-GitText -Context $Context -Arguments @('rev-parse', '--absolute-git-dir')
    $Context.GitDirectory = Resolve-LocalPath -Value $gitDirectory
    $common = Get-GitText -Context $Context -Arguments @('rev-parse', '--git-common-dir')
    if (-not [System.IO.Path]::IsPathRooted($common)) {
        $common = [System.IO.Path]::Combine($Context.Root, $common)
    }
    $Context.CommonDirectory = Resolve-LocalPath -Value $common
    Assert-RepositoryIdle -Context $Context
}

function Assert-RepositoryIdle {
    [CmdletBinding()]
    param([object] $Context)
    foreach ($name in @('MERGE_HEAD', 'CHERRY_PICK_HEAD', 'REVERT_HEAD', 'REBASE_HEAD', 'rebase-merge', 'rebase-apply', 'sequencer')) {
        if (Test-Path -LiteralPath ([System.IO.Path]::Combine($Context.GitDirectory, $name))) {
            Stop-Transfer -Id 'GitOperationInProgress' -Message 'Finish the existing Git merge/rebase/cherry-pick/revert first.'
        }
    }
    $unmerged = Get-GitText -Context $Context -Arguments @('ls-files', '--unmerged', '-z')
    if ($unmerged.Length -ne 0) {
        Stop-Transfer -Id 'GitConflict' -Message 'Resolve index conflicts first.'
    }
}

function Get-BranchReference {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Context)
    $reference = Get-GitText -Context $Context -Arguments @('symbolic-ref', '--quiet', 'HEAD')
    if (-not $reference.StartsWith('refs/heads/', [System.StringComparison]::Ordinal)) {
        Stop-Transfer -Id 'BranchRequired' -Message 'Use an existing checked-out branch, not detached HEAD.'
    }
    return $reference
}

function Get-PushDestination {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Context, [string] $Name)
    $urls = @( (Get-GitText -Context $Context -Arguments @('remote', 'get-url', '--push', '--all', $Name)).Split([char]10) )
    if ($urls.Count -ne 1 -or [string]::IsNullOrWhiteSpace($urls[0])) {
        Stop-Transfer -Id 'PushDestination' -Message 'Exactly one configured push URL is required.' -Code 2
    }
    $url = $urls[0].TrimEnd([char]13)
    if ($url -match '^[A-Za-z0-9+.-]+::' -or $url.IndexOf([char]0) -ge 0) {
        Stop-Transfer -Id 'RemoteHelper' -Message 'Arbitrary remote-helper command destinations are not supported.' -Code 2
    }
    Assert-PushUrl -Url $url
    $mirror = Invoke-TransferGit -Context $Context -Arguments @('config', '--bool', '--get', ('remote.' + $Name + '.mirror')) -AcceptedCodes @(0, 1)
    if ($mirror.Text.Trim() -eq 'true') {
        Stop-Transfer -Id 'MirrorRemote' -Message 'A mirror remote is not suitable for a single-branch transfer.' -Code 2
    }
    return $url
}

function Assert-PushUrl {
    [CmdletBinding()]
    param([string] $Url)
    if ($Url -notmatch '^(https?|ssh)://') {
        return
    }
    $uri = [uri]::new($Url, [System.UriKind]::Absolute)
    if ($uri.Scheme -eq 'ssh') {
        if ($uri.UserInfo.Contains(':')) {
            Stop-Transfer -Id 'CredentialUrl' -Message 'Do not embed a password in the SSH URL.' -Code 2
        }
        return
    }
    if ($uri.UserInfo.Length -gt 0 -or $uri.Query.Length -gt 0 -or $uri.Fragment.Length -gt 0) {
        Stop-Transfer -Id 'CredentialUrl' -Message 'Use an approved credential helper and a plain HTTP repository URL without userinfo, query, or fragment.' -Code 2
    }
}

function Assert-ExternalPlanPath {
    [CmdletBinding()]
    param([object] $Context, [string] $Path)
    foreach ($root in @($Context.Root, $Context.GitDirectory, $Context.CommonDirectory)) {
        if (Test-PathWithin -Path $Path -Directory $root) {
            Stop-Transfer -Id 'PlanInsideRepository' -Message 'Keep plan, checkpoint, and logs outside the worktree and Git metadata.' -Code 2
        }
    }
    Assert-OwnedFilePath -Path $Path
    Assert-OwnedFilePath -Path ($Path + '.state.json')
    Assert-ControlledDirectory -Path ([System.IO.Path]::GetDirectoryName($Path))
}

function Assert-ControlledDirectory {
    [CmdletBinding()]
    param([string] $Path)
    $current = $Path
    while (-not [string]::IsNullOrEmpty($current)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
            if (-not $item.PSIsContainer) {
                Stop-Transfer -Id 'DataDirectory' -Message 'A runtime data parent must be a directory.' -Code 2
            }
            $link = $item.PSObject.Properties['LinkType']
            if ($null -ne $link) {
                if (@('SymbolicLink', 'Junction') -contains $link.Value) {
                    Stop-Transfer -Id 'DataDirectoryLink' -Message 'Use real directory paths for plan/state/logs; directory-link redirection is not supported.' -Code 2
                }
            }
        }
        $parent = [System.IO.Directory]::GetParent($current)
        if ($null -eq $parent) {
            return
        }
        $current = $parent.FullName
    }
}

function Enter-RepositoryLock {
    [CmdletBinding()]
    [OutputType([System.IO.FileStream])]
    param([object] $Context)
    $path = [System.IO.Path]::Combine($Context.CommonDirectory, 'sendshards.lock')
    Assert-OwnedFilePath -Path $path
    try {
        return [System.IO.FileStream]::new($path, [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None, 4096, [System.IO.FileOptions]::DeleteOnClose)
    } catch {
        Stop-Transfer -Id 'RepositoryBusy' -Message 'Cannot acquire the exclusive sendshards repository lock.'
    }
}

function Get-TreeBlob {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Context, [string] $Tree, [string] $RelativePath)
    $text = Get-GitText -Context $Context -Arguments @('ls-tree', '-z', $Tree, '--', $RelativePath)
    if ($text.Length -eq 0) {
        return ''
    }
    $pattern = '^(100644|100755) blob ([0-9a-f]{40}|[0-9a-f]{64})\t' + [regex]::Escape($RelativePath) + '\x00$'
    if ($text -cnotmatch $pattern) {
        Stop-Transfer -Id 'TreeEntry' -Message 'A planned Git path is not one ordinary blob.' -Code 2
    }
    return $Matches[2]
}

function Get-SourceRecords {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([object] $Context, [string] $Directory)
    if (-not (Test-PathWithin -Path $Directory -Directory $Context.Root)) {
        Stop-Transfer -Id 'SourceOutsideRepository' -Message 'The shard folder must be inside the resolved worktree.' -Code 2
    }
    $files = @(Get-ChildItem -LiteralPath $Directory -File -Force -ErrorAction Stop |
        Where-Object {
            $_.Name -match '^.+\.7z\.[0-9]{3,}\.txt$|^.+\.7z\.manifest\.json$'
        })
    if ($files.Count -gt 10000) {
        Stop-Transfer -Id 'FileCount' -Message 'The input exceeds 10,000 shard/manifest files.' -Code 2
    }
    $records = @()
    $prefixLength = $Context.Root.TrimEnd([char[]]'\/').Length + 1
    foreach ($file in $files) {
        $kind = 'Manifest'
        $index = [long]::MaxValue
        $base = ''
        if ($file.Name -match '^(?<base>.+\.7z)\.(?<number>[0-9]{3,})\.txt$') {
            $kind = 'Shard'
            $base = $Matches['base']
            $index = 0L
            if (-not [long]::TryParse($Matches['number'], [ref]$index)) {
                Stop-Transfer -Id 'ShardIndex' -Message 'A shard index is invalid.' -Code 2
            }
            if ($file.Name -cne ($base + '.' + $index.ToString('D3', [System.Globalization.CultureInfo]::InvariantCulture) + '.txt')) {
                Stop-Transfer -Id 'ShardIndex' -Message 'Shard numbers must use canonical numbering.' -Code 2
            }
        } else {
            $base = $file.Name.Substring(0, $file.Name.Length - '.manifest.json'.Length)
        }
        $fingerprint = Get-FileFingerprint -Path $file.FullName -MaximumBytes 268435456
        if ($fingerprint.Bytes -eq 0) {
            Stop-Transfer -Id 'EmptyShard' -Message 'Shard/manifest files must not be empty.' -Code 2
        }
        $relative = $file.FullName.Substring($prefixLength).Replace('\', '/')
        $blob = Get-GitText -Context $Context -Arguments @('hash-object', '--no-filters', '--', $relative)
        $records += [pscustomobject]@{
            Path = $relative
            Kind = $kind
            Base = $base
            Index = $index
            Bytes = $fingerprint.Bytes
            Sha256 = $fingerprint.Sha256
            Blob = $blob
        }
    }
    Assert-ShardGroups -Context $Context -Records $records
    return $records | Sort-Object -Property Base, Index
}

function Assert-ShardGroups {
    [CmdletBinding()]
    param([object] $Context, [object[]] $Records)
    $groups = @($Records | Group-Object -Property Base)
    if ($groups.Count -eq 0) {
        Stop-Transfer -Id 'NoShards' -Message 'No archive shards were found.' -Code 2
    }
    foreach ($group in $groups) {
        $shards = @($group.Group | Where-Object {
                $_.Kind -eq 'Shard'
            } | Sort-Object -Property Index)
        if ($shards.Count -eq 0) {
            Stop-Transfer -Id 'NoShards' -Message 'A manifest has no matching shards.' -Code 2
        }
        for ($index = 0; $index -lt $shards.Count; $index++) {
            if ($shards[$index].Index -ne $index + 1) {
                Stop-Transfer -Id 'ShardSequence' -Message 'Shard numbering is duplicated or not contiguous from 001.' -Code 2
            }
        }
        $manifests = @($group.Group | Where-Object {
                $_.Kind -eq 'Manifest'
            })
        if ($manifests.Count -eq 0) {
            Write-Warning 'A legacy shard set has no checksum manifest.'
            continue
        }
        if ($manifests.Count -ne 1) {
            Stop-Transfer -Id 'ManifestCount' -Message 'A shard set has ambiguous manifests.' -Code 2
        }
        Assert-SourceManifest -Context $Context -Record $manifests[0] -Shards $shards
    }
}

function Assert-SourceManifest {
    [CmdletBinding()]
    param([object] $Context, [object] $Record, [object[]] $Shards)
    $path = [System.IO.Path]::Combine($Context.Root, $Record.Path.Replace('/', '\'))
    $document = Read-TransferJson -Path $path
    Assert-ObjectSchema -Value $document -Keys @('Version', 'Archive', 'Generation', 'ShardCount', 'Shards')
    Assert-IntegerValue -Value $document.Version -Minimum 1 -Maximum 1
    Assert-IntegerValue -Value $document.ShardCount -Minimum 1 -Maximum 10000
    Assert-TextPattern -Value $document.Generation -Pattern '^[0-9a-f]{32}$'
    if ($document.Archive -cne $Record.Base -or $document.Shards -isnot [array]) {
        Stop-Transfer -Id 'ManifestSchema' -Message 'The checksum manifest does not describe this set.' -Code 2
    }
    if ($document.ShardCount -ne $Shards.Count -or $document.Shards.Count -ne $Shards.Count) {
        Stop-Transfer -Id 'ManifestCount' -Message 'The manifest count does not match available shards.' -Code 2
    }
    for ($index = 0; $index -lt $Shards.Count; $index++) {
        $entry = $document.Shards[$index]
        Assert-ObjectSchema -Value $entry -Keys @('Index', 'Name', 'EncodedBytes', 'VolumeBytes', 'Hash', 'VolumeHash')
        Assert-IntegerValue -Value $entry.Index -Minimum 1 -Maximum 10000
        Assert-IntegerValue -Value $entry.EncodedBytes -Minimum 1 -Maximum 268435456
        Assert-IntegerValue -Value $entry.VolumeBytes -Minimum 1 -Maximum 268435456
        Assert-TextPattern -Value $entry.VolumeHash -Pattern '^[0-9a-f]{64}$'
        Assert-ManifestMatch -Entry $entry -Record $Shards[$index]
    }
}

function Assert-ManifestMatch {
    [CmdletBinding()]
    param([object] $Entry, [object] $Record)
    $name = [System.IO.Path]::GetFileName($Record.Path)
    if ($Entry.Index -ne $Record.Index -or $Entry.Name -cne $name) {
        Stop-Transfer -Id 'ManifestSequence' -Message 'A manifest entry does not match shard numbering.' -Code 2
    }
    if ($Entry.EncodedBytes -ne $Record.Bytes -or $Entry.Hash -cne $Record.Sha256) {
        Stop-Transfer -Id 'ShardChecksum' -Message 'A shard does not match the source checksum manifest.' -Code 2
    }
}

function New-TransferPlan {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([object] $Context, [string] $Directory, [string] $RemoteName, [int] $Interval)
    $staged = Get-GitText -Context $Context -Arguments @('diff', '--cached', '--name-only', '-z', '--')
    if ($staged.Length -ne 0) {
        Stop-Transfer -Id 'ForeignIndex' -Message 'Start planning with an empty Git index.'
    }
    $base = Get-GitText -Context $Context -Arguments @('rev-parse', '--verify', 'HEAD')
    $records = @(Get-SourceRecords -Context $Context -Directory $Directory)
    $files = @()
    $pending = 0
    foreach ($record in $records) {
        $isPending = (Get-TreeBlob -Context $Context -Tree $base -RelativePath $record.Path) -cne $record.Blob
        $number = 0
        $wait = 0
        if ($isPending) {
            $pending++
            $number = $pending
            if ($pending -gt 1) {
                $wait = $Interval
            }
        }
        $files += [pscustomobject]@{
            Path = $record.Path
            Kind = $record.Kind
            Bytes = $record.Bytes
            Sha256 = $record.Sha256
            Blob = $record.Blob
            Pending = $isPending
            Item = $number
            WaitSeconds = $wait
        }
    }
    $waitBudget = [math]::Max(0, $pending - 1) * $Interval
    if ($waitBudget -gt 3000) {
        Stop-Transfer -Id 'ScheduleTooLong' -Message 'The planned waits exceed 50 minutes; this rate and count cannot fit the work budget.' -Code 2
    }
    $url = Get-PushDestination -Context $Context -Name $RemoteName
    return [pscustomobject]@{
        SchemaVersion = 1
        PlanId = [guid]::NewGuid().ToString('N')
        CreatedUtc = [datetime]::UtcNow.ToString('o')

        Repository = $Context.Root
        SourceFolder = $Directory
        BranchRef = Get-BranchReference -Context $Context

        Remote = $RemoteName
        PushDestinationSha256 = Get-TextHash -Text $url
        BaseCommit = $base

        BudgetSeconds = 3600
        IntervalSeconds = $Interval
        GitTimeoutSeconds = $Context.GitTimeoutSeconds

        WaitBudgetSeconds = $waitBudget
        PendingCount = $pending
        Files = @($files)
    }
}

function Assert-PlanFileRecord {
    [CmdletBinding()]
    param([object] $Record, [string] $SourceRelative, [int] $PendingNumber, [int] $Interval)
    Assert-ObjectSchema -Value $Record -Keys @('Path', 'Kind', 'Bytes', 'Sha256', 'Blob', 'Pending', 'Item', 'WaitSeconds')
    Assert-PlannedPath -Record $Record -SourceRelative $SourceRelative
    Assert-TextPattern -Value $Record.Sha256 -Pattern '^[0-9a-f]{64}$'
    Assert-TextPattern -Value $Record.Blob -Pattern '^([0-9a-f]{40}|[0-9a-f]{64})$'
    Assert-IntegerValue -Value $Record.Bytes -Minimum 1 -Maximum 268435456
    Assert-IntegerValue -Value $Record.Item -Minimum 0 -Maximum 10000
    Assert-IntegerValue -Value $Record.WaitSeconds -Minimum 0 -Maximum 600
    if ($Record.Pending -isnot [bool]) {
        Stop-Transfer -Id 'PlanBoolean' -Message 'A pending flag must be a JSON boolean.' -Code 2
    }
    $expectedWait = 0
    $expectedItem = 0
    if ($Record.Pending) {
        $expectedItem = $PendingNumber
        if ($PendingNumber -gt 1) {
            $expectedWait = $Interval
        }
    }
    if ($Record.Item -ne $expectedItem -or $Record.WaitSeconds -ne $expectedWait) {
        Stop-Transfer -Id 'PlanSchedule' -Message 'A plan has an inconsistent item number or fixed wait.' -Code 2
    }
}

function Assert-PlannedPath {
    [CmdletBinding()]
    param([object] $Record, [string] $SourceRelative)
    Assert-TextPattern -Value $Record.Path -Pattern '^[^\\:]+$'
    $parent = ''
    $separator = $Record.Path.LastIndexOf('/')
    if ($separator -ge 0) {
        $parent = $Record.Path.Substring(0, $separator)
    }
    if ($parent -cne $SourceRelative) {
        Stop-Transfer -Id 'PlanContainment' -Message 'A planned file must be directly inside the source folder.' -Code 2
    }
    $name = $Record.Path.Substring($separator + 1)
    Assert-TextPattern -Value $Record.Kind -Pattern '^(Shard|Manifest)$'
    $pattern = '^.+\.7z\.[0-9]{3,}\.txt$'
    if ($Record.Kind -ceq 'Manifest') {
        $pattern = '^.+\.7z\.manifest\.json$'
    }
    Assert-TextPattern -Value $name -Pattern $pattern
}

function Assert-TransferPlan {
    [CmdletBinding()]
    param([object] $Document)
    Assert-ObjectSchema -Value $Document -Keys @('SchemaVersion', 'PlanId', 'CreatedUtc', 'Repository', 'SourceFolder', 'BranchRef',
        'Remote', 'PushDestinationSha256', 'BaseCommit', 'BudgetSeconds', 'IntervalSeconds', 'GitTimeoutSeconds', 'WaitBudgetSeconds', 'PendingCount', 'Files')
    Assert-IntegerValue -Value $Document.SchemaVersion -Minimum 1 -Maximum 1
    Assert-IntegerValue -Value $Document.BudgetSeconds -Minimum 3600 -Maximum 3600
    Assert-IntegerValue -Value $Document.IntervalSeconds -Minimum 5 -Maximum 600
    Assert-IntegerValue -Value $Document.GitTimeoutSeconds -Minimum 10 -Maximum 300
    Assert-IntegerValue -Value $Document.WaitBudgetSeconds -Minimum 0 -Maximum 3000
    Assert-IntegerValue -Value $Document.PendingCount -Minimum 0 -Maximum 10000
    foreach ($key in @('Repository', 'SourceFolder', 'BranchRef', 'Remote', 'CreatedUtc')) {
        if ($Document.$key -isnot [string] -or [string]::IsNullOrWhiteSpace($Document.$key)) {
            Stop-Transfer -Id 'PlanString' -Message 'A required plan string is missing.' -Code 2
        }
    }
    Assert-TextPattern -Value $Document.PlanId -Pattern '^[0-9a-f]{32}$'
    Assert-TextPattern -Value $Document.PushDestinationSha256 -Pattern '^[0-9a-f]{64}$'
    Assert-TextPattern -Value $Document.BaseCommit -Pattern '^([0-9a-f]{40}|[0-9a-f]{64})$'
    Assert-TextPattern -Value $Document.Remote -Pattern '^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$'
    $null = ConvertFrom-UtcStamp -Value $Document.CreatedUtc
    if ($Document.Files -isnot [array] -or $Document.Files.Count -lt 1 -or $Document.Files.Count -gt 10000) {
        Stop-Transfer -Id 'PlanFiles' -Message 'The plan file collection is invalid.' -Code 2
    }
    $root = Resolve-LocalPath -Value $Document.Repository
    $source = Resolve-LocalPath -Value $Document.SourceFolder
    Assert-CanonicalPlanPaths -Document $Document -Root $root -Source $source
    if (-not (Test-PathWithin -Path $source -Directory $root)) {
        Stop-Transfer -Id 'PlanContainment' -Message 'The plan source must be inside its repository.' -Code 2
    }
    $relative = $source.Substring($root.TrimEnd([char[]]'\/').Length).TrimStart([char[]]'\/').Replace('\', '/')
    Assert-PlanRecords -Document $Document -SourceRelative $relative
}

function Assert-CanonicalPlanPaths {
    [CmdletBinding()]
    param([object] $Document, [string] $Root, [string] $Source)
    if (-not $Root.Equals($Document.Repository, [System.StringComparison]::OrdinalIgnoreCase)) {
        Stop-Transfer -Id 'PlanAbsolutePath' -Message 'Plan repository paths must be absolute and canonical.' -Code 2
    }
    if (-not $Source.Equals($Document.SourceFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
        Stop-Transfer -Id 'PlanAbsolutePath' -Message 'Plan source paths must be absolute and canonical.' -Code 2
    }
}

function Assert-PlanRecords {
    [CmdletBinding()]
    param([object] $Document, [string] $SourceRelative)
    $names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $pending = 0
    $waits = 0L
    foreach ($record in $Document.Files) {
        Assert-PlanFileRecord -Record $record -SourceRelative $SourceRelative -PendingNumber ($pending + 1) -Interval $Document.IntervalSeconds
        if ($record.Pending -eq $true) {
            $pending++
        }
        if (-not $names.Add($record.Path)) {
            Stop-Transfer -Id 'PlanDuplicatePath' -Message 'A plan repeats a file path.' -Code 2
        }
        $waits += $record.WaitSeconds
    }
    if ($pending -ne $Document.PendingCount -or $waits -ne $Document.WaitBudgetSeconds) {
        Stop-Transfer -Id 'PlanCounts' -Message 'Plan counts and schedule totals are inconsistent.' -Code 2
    }
}

function Assert-PlanRepository {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Context, [object] $Document)
    if (-not $Context.Root.Equals($Document.Repository, [System.StringComparison]::OrdinalIgnoreCase)) {
        Stop-Transfer -Id 'WrongRepository' -Message 'The plan resolves to a different repository.' -Code 2
    }
    if ((Get-BranchReference -Context $Context) -cne $Document.BranchRef) {
        Stop-Transfer -Id 'ChangedBranch' -Message 'The current branch differs from the plan.'
    }
    $url = Get-PushDestination -Context $Context -Name $Document.Remote
    if ((Get-TextHash -Text $url) -cne $Document.PushDestinationSha256) {
        Stop-Transfer -Id 'ChangedDestination' -Message 'The configured push destination changed after planning.'
    }
    return $url
}

function Assert-SourceSnapshot {
    [CmdletBinding()]
    param([object] $Context, [object] $Document)
    $current = @(Get-SourceRecords -Context $Context -Directory $Document.SourceFolder)
    if ($current.Count -ne $Document.Files.Count) {
        Stop-Transfer -Id 'ChangedSet' -Message 'The source shard/manifest inventory differs from the plan.'
    }
    for ($index = 0; $index -lt $current.Count; $index++) {
        $record = $current[$index]
        $expected = $Document.Files[$index]
        if ($record.Path -cne $expected.Path -or $record.Sha256 -cne $expected.Sha256) {
            Stop-Transfer -Id 'ChangedSource' -Message 'A source file differs from the planned order or checksum.'
        }
        if ($record.Bytes -ne $expected.Bytes -or $record.Blob -cne $expected.Blob) {
            Stop-Transfer -Id 'ChangedSource' -Message 'A source file differs from the planned bytes or Git object.'
        }
    }
}

function Get-ItemCommitMessage {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Document, [object] $Record)
    return 'sendshards ' + $Document.PlanId + ' item ' + $Record.Item + ' sha256 ' + $Record.Sha256
}

function Assert-ItemCommit {
    [CmdletBinding()]
    param([object] $Context, [object] $Document, [object] $Record, [string] $Parent, [string] $Commit)
    $parents = Get-GitText -Context $Context -Arguments @('rev-list', '--parents', '-n', '1', $Commit)
    if ($parents -cne ($Commit + ' ' + $Parent)) {
        Stop-Transfer -Id 'UnexpectedParent' -Message 'A transfer commit has an unexpected parent or merge.'
    }
    $message = Get-GitText -Context $Context -Arguments @('show', '-s', '--format=%B', $Commit)
    if ($message -cne (Get-ItemCommitMessage -Document $Document -Record $Record)) {
        Stop-Transfer -Id 'ForeignCommit' -Message 'A commit is not an exact planned transfer commit.'
    }
    $changes = Get-GitText -Context $Context -Arguments @('diff-tree', '--no-commit-id', '--name-only', '-r', '-z', $Parent, $Commit, '--')
    if ($changes -cne ($Record.Path + [char]0)) {
        Stop-Transfer -Id 'CommitContents' -Message 'A commit includes missing or unrelated changes.'
    }
    if ((Get-TreeBlob -Context $Context -Tree $Commit -RelativePath $Record.Path) -cne $Record.Blob) {
        Stop-Transfer -Id 'CommitBlob' -Message 'Committed shard bytes differ from the plan.'
    }
}

function Get-TransferChain {
    [CmdletBinding()]
    [OutputType([string[]])]
    param([object] $Context, [object] $Document)
    $head = Get-GitText -Context $Context -Arguments @('rev-parse', '--verify', 'HEAD')
    $null = Invoke-TransferGit -Context $Context -Arguments @('merge-base', '--is-ancestor', $Document.BaseCommit, $head)
    $text = Get-GitText -Context $Context -Arguments @('rev-list', '--reverse', ($Document.BaseCommit + '..' + $head))
    $commits = @()
    if ($text.Length -gt 0) {
        $commits = @($text.Split([char]10))
    }
    $pending = @($Document.Files | Where-Object {
            $_.Pending
        })
    if ($commits.Count -gt $pending.Count) {
        Stop-Transfer -Id 'ForeignHistory' -Message 'The branch contains commits beyond this plan.'
    }
    $parent = $Document.BaseCommit
    for ($index = 0; $index -lt $commits.Count; $index++) {
        Assert-ItemCommit -Context $Context -Document $Document -Record $pending[$index] -Parent $parent -Commit $commits[$index]
        $parent = $commits[$index]
    }
    foreach ($record in @($Document.Files | Where-Object {
            -not $_.Pending
        })) {
        if ((Get-TreeBlob -Context $Context -Tree $Document.BaseCommit -RelativePath $record.Path) -cne $record.Blob) {
            Stop-Transfer -Id 'InvalidSkip' -Message 'A skipped file is not present in the planned base commit.' -Code 2
        }
    }
    return ,$commits
}

function Get-RemoteCommit {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Context, [string] $Url, [string] $Branch)
    $text = Get-GitText -Context $Context -Arguments @('ls-remote', '--refs', '--exit-code', '--', $Url, $Branch)
    $pattern = '^([0-9a-f]{40}|[0-9a-f]{64})\t' + [regex]::Escape($Branch) + '$'
    if ($text -cnotmatch $pattern) {
        Stop-Transfer -Id 'RemoteReference' -Message 'The target branch must already exist with one exact remote reference.'
    }
    return $Matches[1]
}

function ConvertFrom-UtcStamp {
    [CmdletBinding()]
    [OutputType([datetime])]
    param([object] $Value)
    if ($Value -isnot [string] -or $Value -cnotmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{7}Z$') {
        Stop-Transfer -Id 'UtcTimestamp' -Message 'A timestamp must be an invariant UTC round-trip value.' -Code 2
    }
    return [datetime]::ParseExact($Value, 'o', [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::RoundtripKind)
}

function Read-TransferState {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([object] $Context, [object] $Document, [string] $Path, [string] $PlanHash)
    if (-not [System.IO.File]::Exists($Path)) {
        return [pscustomobject]@{
            SchemaVersion = 1
            PlanId = $Document.PlanId
            PlanHash = $PlanHash

            StartedUtc = $Context.StartedUtc.ToString('o')
            LastAttemptUtc = $null
            LastSuccessUtc = $null

            VerifiedCount = 0
            VerifiedCommit = $Document.BaseCommit
            Status = 'Running'
        }
    }
    $state = Read-TransferJson -Path $Path
    Assert-ObjectSchema -Value $state -Keys @('SchemaVersion', 'PlanId', 'PlanHash', 'StartedUtc', 'LastAttemptUtc', 'LastSuccessUtc', 'VerifiedCount', 'VerifiedCommit', 'Status')
    Assert-IntegerValue -Value $state.SchemaVersion -Minimum 1 -Maximum 1
    Assert-IntegerValue -Value $state.VerifiedCount -Minimum 0 -Maximum $Document.PendingCount
    Assert-TextPattern -Value $state.VerifiedCommit -Pattern '^([0-9a-f]{40}|[0-9a-f]{64})$'
    if ($state.PlanId -cne $Document.PlanId -or $state.PlanHash -cne $PlanHash) {
        Stop-Transfer -Id 'ChangedPlan' -Message 'The plan bytes differ from the checkpoint; do not edit an active plan.' -Code 2
    }
    if (@('Running', 'Paused', 'Complete') -cnotcontains $state.Status) {
        Stop-Transfer -Id 'StateStatus' -Message 'The checkpoint status is invalid.' -Code 2
    }
    if ($state.Status -ceq 'Complete' -and $state.VerifiedCount -ne $Document.PendingCount) {
        Stop-Transfer -Id 'StateCompletion' -Message 'A completed checkpoint has an inconsistent verified count.' -Code 2
    }
    $null = ConvertFrom-UtcStamp -Value $state.StartedUtc
    foreach ($value in @($state.LastAttemptUtc, $state.LastSuccessUtc)) {
        if ($null -ne $value) {
            $null = ConvertFrom-UtcStamp -Value $value
        }
    }
    return $state
}

function Set-TransferDeadline {
    [CmdletBinding()]
    param([object] $Context, [object] $State)
    $start = ConvertFrom-UtcStamp -Value $State.StartedUtc
    if ($start -gt $Context.StartedUtc.AddSeconds(5)) {
        Stop-Transfer -Id 'ClockOrState' -Message 'The checkpoint start time is in the future; verify the clock and state.' -Code 2
    }
    $Context.DeadlineUtc = $start.AddSeconds(3600)
    $Context.AllowedSeconds = ($Context.DeadlineUtc - $Context.StartedUtc).TotalSeconds
}

function Sync-TransferState {
    [CmdletBinding()]
    [OutputType([int])]
    param([object] $Context, [object] $Document, [object] $State, [string[]] $Commits, [string] $Url, [string] $StatePath)
    $remoteCommit = Get-RemoteCommit -Context $Context -Url $Url -Branch $Document.BranchRef
    $positions = @($Document.BaseCommit) + @($Commits)
    if ($State.VerifiedCount -gt $Commits.Count) {
        Stop-Transfer -Id 'MissingHistory' -Message 'Local transfer history is missing previously verified commits.'
    }
    $verified = [array]::IndexOf($positions, $remoteCommit)
    if ($verified -lt 0) {
        Stop-Transfer -Id 'RemoteDivergence' -Message 'The remote branch differs from the base or verified transfer history.'
    }
    if ($verified -lt $State.VerifiedCount) {
        Stop-Transfer -Id 'RemoteRewound' -Message 'The remote branch was rewound after a verified checkpoint.'
    }
    if ($State.VerifiedCommit -cne $positions[$State.VerifiedCount]) {
        Stop-Transfer -Id 'CheckpointHistory' -Message 'The checkpoint does not match local transfer history.'
    }
    if ($Commits.Count -gt $verified + 1) {
        Stop-Transfer -Id 'UnexpectedLocalHistory' -Message 'More than one local transfer commit is awaiting push.'
    }
    if ($verified -gt $State.VerifiedCount) {
        $State.LastSuccessUtc = [datetime]::UtcNow.ToString('o')
    }
    $State.VerifiedCount = $verified
    $State.VerifiedCommit = $remoteCommit
    $State.Status = 'Running'
    if ($verified -eq $Document.PendingCount) {
        $State.Status = 'Complete'
    }
    Write-AtomicJson -Path $StatePath -Document $State
    return $verified
}

function Assert-TransferIndex {
    [CmdletBinding()]
    param([object] $Context, [object] $Record)
    $text = Get-GitText -Context $Context -Arguments @('diff', '--cached', '--name-only', '-z', '--')
    if ($text.Length -eq 0) {
        return
    }
    if ($null -eq $Record) {
        Stop-Transfer -Id 'ForeignIndex' -Message 'The index contains staged changes after the plan is complete.'
    }
    if ($text -cne ($Record.Path + [char]0)) {
        Stop-Transfer -Id 'ForeignIndex' -Message 'The index contains unrelated staged changes.'
    }
    $entry = Get-GitText -Context $Context -Arguments @('ls-files', '--stage', '-z', '--', $Record.Path)
    $pattern = '^(100644|100755) ' + [regex]::Escape($Record.Blob) + ' 0\t' + [regex]::Escape($Record.Path) + '\x00$'
    if ($entry -cnotmatch $pattern) {
        Stop-Transfer -Id 'ChangedIndexBytes' -Message 'The staged file differs from planned raw bytes; inspect Git attributes/filters.'
    }
}

function Wait-TransferInterval {
    [CmdletBinding()]
    param([object] $Context, [object] $State, [int] $Interval)
    $due = $Context.StartedUtc
    foreach ($stamp in @($State.LastAttemptUtc, $State.LastSuccessUtc)) {
        if ($null -ne $stamp) {
            $candidate = (ConvertFrom-UtcStamp -Value $stamp).AddSeconds($Interval)
            if ($candidate -gt $due) {
                $due = $candidate
            }
        }
    }
    $delay = [math]::Max(0, ($due - [datetime]::UtcNow).TotalSeconds)
    if ($delay -gt (Get-RemainingTime -Context $Context)) {
        Stop-Transfer -Id 'Deadline' -Message 'The next permitted push falls outside this plan budget.' -Code 4
    }
    Write-TransferEvent -Context $Context -Event 'Wait' -Fields @{
        waitSeconds = [math]::Ceiling($delay)
    }
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($watch.Elapsed.TotalSeconds -lt $delay) {
        $null = Get-RemainingTime -Context $Context
        Start-Sleep -Milliseconds ([int][math]::Min(1000, [math]::Max(1, ($delay - $watch.Elapsed.TotalSeconds) * 1000)))
    }
}

function New-ItemCommit {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Context, [object] $Document, [object] $Record, [string] $Parent)
    Assert-RepositoryIdle -Context $Context
    Assert-TransferIndex -Context $Context -Record $Record
    $path = [System.IO.Path]::Combine($Context.Root, $Record.Path.Replace('/', '\'))
    $current = Get-FileFingerprint -Path $path -MaximumBytes 268435456
    if ($current.Sha256 -cne $Record.Sha256 -or $current.Bytes -ne $Record.Bytes) {
        Stop-Transfer -Id 'ChangedSource' -Message 'The next shard changed after planning.'
    }
    if ((Get-GitText -Context $Context -Arguments @('rev-parse', '--verify', 'HEAD')) -cne $Parent) {
        Stop-Transfer -Id 'ChangedHead' -Message 'The branch changed during transfer.'
    }
    $null = Invoke-TransferGit -Context $Context -Arguments @('add', '--', $Record.Path)
    Assert-TransferIndex -Context $Context -Record $Record
    $message = Get-ItemCommitMessage -Document $Document -Record $Record
    $null = Invoke-TransferGit -Context $Context -Arguments @('commit', '-m', $message, '--no-status')
    $commit = Get-GitText -Context $Context -Arguments @('rev-parse', '--verify', 'HEAD')
    Assert-ItemCommit -Context $Context -Document $Document -Record $Record -Parent $Parent -Commit $commit
    Assert-TransferIndex -Context $Context -Record $null
    return $commit
}

function Send-ItemCommit {
    [CmdletBinding()]
    param([object] $Context, [object] $Document, [object] $State, [string] $Commit, [string] $Url, [string] $StatePath)
    if ((Get-RemoteCommit -Context $Context -Url $Url -Branch $Document.BranchRef) -cne $State.VerifiedCommit) {
        Stop-Transfer -Id 'RemoteChanged' -Message 'The remote branch changed before this push.'
    }
    $State.LastAttemptUtc = [datetime]::UtcNow.ToString('o')
    Write-AtomicJson -Path $StatePath -Document $State
    Write-TransferEvent -Context $Context -Event 'PushAttempt' -Fields @{
        item = $State.VerifiedCount + 1
    }
    # Exact source OID, one branch refspec, no force, no tag propagation, hooks active.
    $null = Invoke-TransferGit -Context $Context -Arguments @('push', '--porcelain', '--no-follow-tags', '--', $Document.Remote, ($Commit + ':' + $Document.BranchRef))
    $observed = Get-RemoteCommit -Context $Context -Url $Url -Branch $Document.BranchRef
    if ($observed -cne $Commit) {
        Stop-Transfer -Id 'PushUnverified' -Message 'Push completion could not be verified at the remote branch.'
    }
    $State.VerifiedCount++
    $State.VerifiedCommit = $Commit
    $State.LastSuccessUtc = [datetime]::UtcNow.ToString('o')
    if ($State.VerifiedCount -eq $Document.PendingCount) {
        $State.Status = 'Complete'
    }
    Write-AtomicJson -Path $StatePath -Document $State
    Write-TransferEvent -Context $Context -Event 'VerifiedPush' -Fields @{
        verifiedCount = $State.VerifiedCount
        planId = $Document.PlanId
    }
}

function Invoke-TransferPush {
    [CmdletBinding()]
    param([object] $Context, [object] $Document, [string] $PlanFile)
    $planHash = (Get-FileFingerprint -Path $PlanFile -MaximumBytes 8388608).Sha256
    if ($planHash -cne $Context.PlanHash) {
        Stop-Transfer -Id 'ChangedPlan' -Message 'The plan changed after it was read.' -Code 2
    }
    $statePath = $PlanFile + '.state.json'
    $state = Read-TransferState -Context $Context -Document $Document -Path $statePath -PlanHash $planHash
    if ($state.Status -cne 'Complete') {
        Set-TransferDeadline -Context $Context -State $state
    }
    $url = Assert-PlanRepository -Context $Context -Document $Document
    Assert-SourceSnapshot -Context $Context -Document $Document
    $commits = Get-TransferChain -Context $Context -Document $Document
    $pending = @($Document.Files | Where-Object {
            $_.Pending
        })
    # Save the original start before contacting the remote or making Git mutations.
    if (-not [System.IO.File]::Exists($statePath)) {
        Write-AtomicJson -Path $statePath -Document $state -NewOnly
    }
    $verified = Sync-TransferState -Context $Context -Document $Document -State $state -Commits $commits -Url $url -StatePath $statePath
    while ($verified -lt $pending.Count) {
        $record = $pending[$verified]
        Wait-TransferInterval -Context $Context -State $state -Interval $Document.IntervalSeconds
        $url = Assert-PlanRepository -Context $Context -Document $Document
        Assert-TransferIndex -Context $Context -Record $record
        $commit = ''
        if ($commits.Count -gt $verified) {
            $commit = $commits[$verified]
            Assert-TransferIndex -Context $Context -Record $null
        } else {
            $commit = New-ItemCommit -Context $Context -Document $Document -Record $record -Parent $state.VerifiedCommit
            $commits += $commit
        }
        Send-ItemCommit -Context $Context -Document $Document -State $state -Commit $commit -Url $url -StatePath $statePath
        $verified = $state.VerifiedCount
    }
    Write-TransferEvent -Context $Context -Event 'Complete' -Fields @{
        verifiedCount = $verified
        status = 'Complete'
        planId = $Document.PlanId
    }
    Write-Verbose ('All planned files are verified at the configured remote. Checkpoint: ' + $statePath)
}

function Get-VerifiedPlanInput {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $Path)
    $before = (Get-FileFingerprint -Path $Path -MaximumBytes 8388608).Sha256
    $document = Read-TransferJson -Path $Path
    Assert-TransferPlan -Document $document
    $after = (Get-FileFingerprint -Path $Path -MaximumBytes 8388608).Sha256
    if ($before -cne $after) {
        Stop-Transfer -Id 'ChangedPlan' -Message 'The plan changed while it was read.' -Code 2
    }
    return [pscustomobject]@{
        Document = $document
        Sha256 = $before
    }
}

function Assert-InvocationMode {
    [CmdletBinding()]
    param([bool] $IsPlan, [bool] $IsPush)
    if ($IsPlan -eq $IsPush) {
        Stop-Transfer -Id 'Mode' -Message 'Specify exactly one of -Plan or -Push.' -Code 2
    }
}

function Write-TransferFailure {
    [CmdletBinding()]
    [OutputType([int])]
    param([object] $Context, [System.Management.Automation.ErrorRecord] $Failure)
    $id = 'UnexpectedFailure'
    $message = 'An unexpected failure occurred. Use -Debug for sanitized diagnostics; inspect Git and the checkpoint before resuming.'
    $code = Get-TransferFailureCode -Failure $Failure
    if ($code -eq 3) {
        $message = 'A filesystem, dependency, or permission operation failed. Inspect the checkpoint and sanitized diagnostics before resuming.'
    }
    if ($code -eq 2) {
        $message = 'An input or document is invalid. Check the supplied paths and plan format.'
    }
    if ($Failure.Exception.Data.Contains('TransferId')) {
        $id = [string]$Failure.Exception.Data['TransferId']
        $message = $Failure.Exception.Message
        $code = [int]$Failure.Exception.Data['TransferCode']
    }
    if ($null -ne $Context) {
        Write-TransferEvent -Context $Context -Event 'Failure' -Fields @{
            errorId = $id
            errorType = $Failure.Exception.GetType().FullName
            line = $Failure.InvocationInfo.ScriptLineNumber
            status = 'Incomplete'
        }
    }
    Write-Error -Message $message -ErrorId $id -Category OperationStopped -ErrorAction Continue
    return $code
}

function Get-TransferFailureCode {
    [CmdletBinding()]
    [OutputType([int])]
    param([System.Management.Automation.ErrorRecord] $Failure)
    $exception = $Failure.Exception.GetBaseException()
    if ($exception -is [System.IO.IOException] -or $exception -is [System.UnauthorizedAccessException] -or $exception -is [System.ComponentModel.Win32Exception]) {
        return 3
    }
    if ($exception -is [System.ArgumentException] -or $exception -is [System.FormatException] -or $exception -is [System.Runtime.Serialization.SerializationException] -or $exception -is [System.Xml.XmlException]) {
        return 2
    }
    return 1
}

function Invoke-PlanPublication {
    [CmdletBinding()]
    param([object] $Context, [object] $Document, [string] $Path)
    if (Test-Path -LiteralPath $Path) {
        Stop-Transfer -Id 'PlanExists' -Message 'Choose a new plan filename; existing plans are not overwritten.' -Code 2
    }
    if (Test-Path -LiteralPath ($Path + '.state.json')) {
        Stop-Transfer -Id 'CheckpointExists' -Message 'A checkpoint already exists for this filename.' -Code 2
    }
    # Revalidate the frozen view after acquiring the tool lock.
    $null = Assert-PlanRepository -Context $Context -Document $Document
    Assert-SourceSnapshot -Context $Context -Document $Document
    if ((Get-GitText -Context $Context -Arguments @('rev-parse', '--verify', 'HEAD')) -cne $Document.BaseCommit) {
        Stop-Transfer -Id 'ChangedHead' -Message 'The branch changed while planning.'
    }
    Write-AtomicJson -Path $Path -Document $Document -NewOnly
    Write-TransferEvent -Context $Context -Event 'PlanCreated' -Fields @{
        fileCount = $Document.Files.Count
        pendingCount = $Document.PendingCount
        planId = $Document.PlanId
        status = 'Complete'
    }
    Write-Verbose ('Plan created: ' + $Path)
}

# Thin executable boundary: no dot-sourcing interface, success objects, or custom prompts.
$context = $null
$lock = $null
$exitCode = 0
try {
    Assert-InvocationMode -IsPlan $Plan.IsPresent -IsPush $Push.IsPresent
    Assert-HostPrerequisite
    $planFile = Resolve-LocalPath -Value $PlanPath
    if ($Plan) {
        $source = Resolve-LocalPath -Value $SourcePath
        $context = New-TransferContext -WorkingPath $source -ExecutablePath $GitPath -Timeout $GitTimeoutSeconds -EnableDebug $debugEnabled
        Set-RepositoryContext -Context $context
        Assert-ExternalPlanPath -Context $context -Path $planFile
        $document = New-TransferPlan -Context $context -Directory $source -RemoteName $Remote -Interval $WaitSeconds
        Assert-TransferPlan -Document $document
        $action = 'Create the fixed-rate Git shard transfer plan'
    } else {
        # JSON reader assembly is loaded here before reading the input document.
        Add-Type -AssemblyName System.Runtime.Serialization -ErrorAction Stop
        $inputPlan = Get-VerifiedPlanInput -Path $planFile
        $document = $inputPlan.Document
        $context = New-TransferContext -WorkingPath $document.Repository -ExecutablePath $GitPath -Timeout $document.GitTimeoutSeconds -EnableDebug $debugEnabled
        $context.PlanHash = $inputPlan.Sha256
        $context.DeadlineUtc = $context.StartedUtc.AddSeconds(3600)
        Set-RepositoryContext -Context $context
        Assert-ExternalPlanPath -Context $context -Path $planFile
        $null = Assert-PlanRepository -Context $context -Document $document
        $action = 'Execute or resume the fixed-rate Git shard transfer plan'
    }
    if ($PSCmdlet.ShouldProcess($planFile, $action)) {
        $lock = Enter-RepositoryLock -Context $context
        $directory = [System.IO.Path]::GetDirectoryName($planFile)
        $null = [System.IO.Directory]::CreateDirectory($directory)
        Initialize-TransferLog -Context $context -Directory $directory
        Write-TransferEvent -Context $context -Event 'Request' -Fields @{
            operation = $PSCmdlet.ParameterSetName
            planId = $document.PlanId
            pendingCount = $document.PendingCount
            waitSeconds = $document.IntervalSeconds
        }
        if ($Plan) {
            Invoke-PlanPublication -Context $context -Document $document -Path $planFile
        }
        else {
            Invoke-TransferPush -Context $context -Document $document -PlanFile $planFile
        }
    }
} catch {
    $exitCode = Write-TransferFailure -Context $context -Failure $_
} finally {
    if ($null -ne $lock) {
        $lock.Dispose()
    }
    if ($null -ne $context) {
        $context.Watch.Stop()
    }
}
exit $exitCode
