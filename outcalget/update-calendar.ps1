#requires -Version 5.1

# .EXTERNALHELP ../docs/outcalget/update-calendar-help.xml
[CmdletBinding(PositionalBinding = $false, SupportsShouldProcess = $true, ConfirmImpact = 'Low', DefaultParameterSetName = 'Profile')]
param(
    [string] $ExporterPath = (Join-Path -Path $PSScriptRoot -ChildPath 'outcalget.ps1'),
    [Parameter(ParameterSetName = 'Config')]
    [string] $Config,
    [Parameter(ParameterSetName = 'Profile')]
    [Alias('Profile')]
    [string] $ExportProfile,
    [switch] $Scheduled,
    [string] $LogDirectory,
    [ValidateRange(15, 1800)]
    [int] $GitTimeoutSeconds = 120
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'
$script:CalendarDebug = $PSBoundParameters.ContainsKey('Debug') -and [bool]$PSBoundParameters['Debug']
$DebugPreference = 'SilentlyContinue'
$script:CalendarRecord = $null
$script:CalendarLog = $null
$updateLock = $null
$updateExit = 0
$updateHandled = $false

function Stop-CalendarUpdate {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Raises a classified exception only; no state mutation.')]
    [CmdletBinding()]
    param([string] $Code, [string] $Message)
    $failure = [InvalidOperationException]::new($Message)
    $failure.Data['CalendarCode'] = $Code
    throw $failure
}

function Get-PublishedCalendarPath {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Result)
    foreach ($key in @('resultSchemaVersion', 'operation', 'exitCode', 'published', 'outputPath')) {
        if ($null -eq $result.PSObject.Properties[$key]) {
            Stop-CalendarUpdate -Code 'EXPORT_RESULT_INVALID' -Message 'outcalget.ps1 returned an incomplete export result. No commit or push was attempted.'
        }
    }
    if (($result.resultSchemaVersion -isnot [int] -and $result.resultSchemaVersion -isnot [long]) -or $result.resultSchemaVersion -ne 1 -or
        $result.operation -isnot [string] -or $result.operation -cne 'export' -or
        ($result.exitCode -isnot [int] -and $result.exitCode -isnot [long]) -or $result.exitCode -ne 0 -or
        $result.published -isnot [bool] -or -not $result.published -or
        $result.outputPath -isnot [string] -or [string]::IsNullOrWhiteSpace($result.outputPath)) {
        Stop-CalendarUpdate -Code 'EXPORT_RESULT_INVALID' -Message 'outcalget.ps1 did not confirm a successfully published calendar. No commit or push was attempted.'
    }
    if (-not [System.IO.Path]::IsPathRooted($result.outputPath) -or
        -not (Test-Path -LiteralPath $result.outputPath -PathType Leaf)) {
        Stop-CalendarUpdate -Code 'EXPORT_RESULT_INVALID' -Message 'The published calendar path is not an existing absolute file. No commit or push was attempted.'
    }
    return (Get-Item -LiteralPath $result.outputPath).FullName
}

function Get-CalendarExportMessage {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Detail)
    if ($Detail -isnot [pscustomobject]) {
        return
    }
    $code = $Detail.PSObject.Properties['code']
    $message = $Detail.PSObject.Properties['message']
    if ($null -ne $code -and $code.Value -is [string] -and $null -ne $message -and $message.Value -is [string]) {
        return ('[{0}]: {1}' -f $code.Value, $message.Value)
    }
}

function Write-CalendarExportNotice {
    [CmdletBinding()]
    param([object] $Result)
    $record = $Result.PSObject.Properties['runLogPath']
    if ($null -ne $record -and $record.Value -is [string] -and -not [string]::IsNullOrWhiteSpace($record.Value)) {
        Write-Information -InformationAction Continue -MessageData ('Outlook export run record: ' + $record.Value)
    }
    $warnings = $Result.PSObject.Properties['warnings']
    if ($null -eq $warnings -or $warnings.Value -isnot [array]) {
        return
    }
    foreach ($warning in $warnings.Value) {
        $message = Get-CalendarExportMessage -Detail $warning
        if (-not [string]::IsNullOrEmpty($message)) {
            Write-Warning -Message $message
        }
    }
}

function Read-CalendarExportResult {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $Text, [int] $ExitCode)
    try {
        $result = ConvertFrom-Json -InputObject $Text -ErrorAction Stop
    } catch {
        Stop-CalendarUpdate -Code 'EXPORT_RESULT_INVALID' -Message "outcalget.ps1 returned no valid JSON export result (exit code $ExitCode). No commit or push was attempted."
    }
    if ($result -isnot [pscustomobject]) {
        Stop-CalendarUpdate -Code 'EXPORT_RESULT_INVALID' -Message "outcalget.ps1 did not return an export result (exit code $ExitCode). No commit or push was attempted."
    }
    $script:CalendarRecord.exporterExitCode = $ExitCode
    $runId = [guid]::Empty
    if ($null -ne $result.PSObject.Properties['runId'] -and
        [guid]::TryParse([string]$result.runId, [ref]$runId)) {
        $script:CalendarRecord.exporterRunId = $runId.ToString('D')
    }
    if ($null -ne $result.PSObject.Properties['published'] -and $result.published -is [bool]) {
        $script:CalendarRecord.exportPublished = $result.published
    }
    Write-CalendarExportNotice -Result $result
    if ($ExitCode -ne 0) {
        $failure = "outcalget.ps1 failed (exit code $ExitCode)."
        $errorProperty = $result.PSObject.Properties['error']
        if ($null -ne $errorProperty) {
            $message = Get-CalendarExportMessage -Detail $errorProperty.Value
            if (-not [string]::IsNullOrEmpty($message)) {
                $failure += ' ' + $message
            }
        }
        $published = $result.PSObject.Properties['published']
        if ($null -ne $published -and $published.Value -is [bool] -and $published.Value) {
            $failure += ' Output was published before the later failure; inspect the run record before retrying.'
        }
        Stop-CalendarUpdate -Code 'EXPORT_FAILED' -Message ($failure + ' No commit or push was attempted.')
    }
    return $result
}

function Get-CalendarRepository {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $GitPath, [string] $CalendarPath)
    $directory = [System.IO.Path]::GetDirectoryName($CalendarPath)
    $rootOutput = Invoke-CalendarGit -GitPath $GitPath -Arguments @('-C', $directory, 'rev-parse', '--show-cdup')
    $root = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($directory, $rootOutput.Text.TrimEnd([char[]]"`r`n")))
    $prefix = $root.TrimEnd([char[]]'\/') + [System.IO.Path]::DirectorySeparatorChar
    if (-not $CalendarPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'The published calendar is outside the resolved Git working tree.'
    }
    return [pscustomobject]@{
        Root = $root
        RelativePath = $CalendarPath.Substring($prefix.Length).Replace('\', '/')
    }
}

function Write-CalendarRecord {
    param([string] $Stage)
    if ($Stage) { $script:CalendarRecord.stage = $Stage }
    $script:CalendarRecord.updatedUtc = [datetime]::UtcNow.ToString('o')
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($script:CalendarRecord | ConvertTo-Json -Depth 4 -Compress))
    $script:CalendarLog.Position = 0
    $script:CalendarLog.Write($bytes, 0, $bytes.Length)
    $script:CalendarLog.SetLength($bytes.Length)
    $script:CalendarLog.Flush($true)
}

function Initialize-CalendarRecord {
    param([string] $Directory)
    [void][IO.Directory]::CreateDirectory($Directory)
    $id = [guid]::NewGuid().ToString('D')
    $path = Join-Path -Path $Directory -ChildPath ('update-calendar-' + [datetime]::UtcNow.ToString('yyyyMMddTHHmmssZ') + '-' + $id + '.json')
    $script:CalendarLog = [IO.File]::Open($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    $script:CalendarRecord = [ordered]@{
        recordSchemaVersion = 1
        tool = 'update-calendar'
        version = '1.1.0'
        runId = $id
        startedUtc = [datetime]::UtcNow.ToString('o')
        updatedUtc = $null
        scheduled = [bool]$Scheduled
        stage = 'initialise'
        completed = $false
        outcome = 'running'
        exitCode = $null
        exporterRunId = $null
        exporterExitCode = $null
        exportPublished = $null
        calendarChanged = $null
        commitCreated = $false
        commit = $null
        pushed = $false
        remoteConfirmed = $false
        gitOutcomeUncertain = $false
        recoveryRequired = $false
        gitExitCode = $null
        errorCode = $null
        diagnostics = $null
    }
    if ($script:CalendarDebug) {
        $script:CalendarRecord.diagnostics = @{
            powershellVersion = $PSVersionTable.PSVersion.ToString()
            scriptSha256 = $null
        }
        try {
            $script:CalendarRecord.diagnostics.scriptSha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant()
        } catch { Write-Warning -Message 'Optional script hash diagnostics are unavailable.' }
    }
    Write-CalendarRecord
    Write-Information -InformationAction Continue -MessageData ('Calendar update run record: ' + $path)
    $closed = 0
    try {
        $files = @(Get-ChildItem -LiteralPath $Directory -Filter 'update-calendar-*.json' -File | Sort-Object LastWriteTimeUtc -Descending)
    } catch {
        Write-Warning -Message 'Old calendar update records could not be listed for retention.'
        return
    }
    foreach ($file in $files) {
        if ($file.FullName -eq $path) { continue }
        $closed++
        if ($closed -le 200 -and $file.LastWriteTimeUtc -ge [datetime]::UtcNow.AddDays(-14)) { continue }
        $guard = $null
        try {
            $guard = [IO.File]::Open($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Delete)
            [IO.File]::Delete($file.FullName)
        } catch {
            Write-Warning -Message 'An old calendar update record could not be pruned.'
        } finally {
            if ($null -ne $guard) { $guard.Dispose() }
        }
    }
}

function ConvertTo-CalendarArgument {
    param([string] $Value)
    return '"' + [regex]::Replace([regex]::Replace($Value, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1') + '"'
}

function Invoke-CalendarGit {
    [CmdletBinding()]
    param([string] $GitPath, [string[]] $Arguments, [int[]] $AllowedExitCodes = @(0))
    $info = [Diagnostics.ProcessStartInfo]::new()
    $info.FileName = $GitPath
    $info.Arguments = (@($Arguments | ForEach-Object { ConvertTo-CalendarArgument -Value $_ }) -join ' ')
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardInput = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    $info.StandardErrorEncoding = [Text.UTF8Encoding]::new($false)
    $info.EnvironmentVariables['GIT_TERMINAL_PROMPT'] = '0'
    $info.EnvironmentVariables['GCM_INTERACTIVE'] = 'Never'
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $info
    $pendingOwned = $false
    $started = $false
    $nativeComplete = $false
    $script:CalendarRecord.gitExitCode = $null
    try {
        $pending = [IO.File]::Open($script:CalendarPendingPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        $pendingOwned = $true
        try {
            $bytes = [Text.UTF8Encoding]::new($false).GetBytes($script:CalendarRecord.runId)
            $pending.Write($bytes, 0, $bytes.Length)
            $pending.Flush($true)
        } finally { $pending.Dispose() }
        [void]$process.Start()
        $started = $true
        $process.StandardInput.Close()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($GitTimeoutSeconds * 1000)) {
            $script:CalendarRecord.gitOutcomeUncertain = $true
            $script:CalendarRecord.recoveryRequired = $true
            try { $process.Kill() } catch { Write-Warning -Message 'The timed-out Git process could not be stopped.' }
            Stop-CalendarUpdate -Code 'GIT_TIMEOUT' -Message 'Git timed out. Inspect Git, hooks and credential/signing helpers before retrying; child work may remain active.'
        }
        if (-not $stdout.Wait(1000) -or -not $stderr.Wait(1000)) {
            $script:CalendarRecord.gitOutcomeUncertain = $true
            $script:CalendarRecord.recoveryRequired = $true
            Stop-CalendarUpdate -Code 'GIT_OUTPUT_OPEN' -Message 'Git output did not close. Inspect native helpers before retrying.'
        }
        $nativeComplete = $true
        $script:CalendarRecord.gitExitCode = $process.ExitCode
        if ($process.ExitCode -notin $AllowedExitCodes) {
            Stop-CalendarUpdate -Code 'GIT_COMMAND_FAILED' -Message ('Git failed during {0} (exit code {1}). Review the run record and repository locally; a commit may remain local.' -f $script:CalendarRecord.stage, $process.ExitCode)
        }
        return [pscustomobject]@{ Text = $stdout.Result; ExitCode = $process.ExitCode }
    } finally {
        if ($pendingOwned -and $started -and -not $nativeComplete) {
            $script:CalendarRecord.gitOutcomeUncertain = $true
            $script:CalendarRecord.recoveryRequired = $true
        }
        $process.Dispose()
        if ($pendingOwned -and ($nativeComplete -or -not $started)) {
            try { [IO.File]::Delete($script:CalendarPendingPath) } catch {
                $script:CalendarRecord.recoveryRequired = $true
                Write-Warning -Message 'The Git recovery marker could not be cleared. Later updates require review.'
            }
        }
    }
}

function Get-CalendarPushTarget {
    param([string] $GitPath, [string] $Root)
    $prefix = @('-C', $Root)
    $branch = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('symbolic-ref', '--quiet', '--short', 'HEAD'))).Text.Trim()
    $remote = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('config', '--get-all', ('branch.' + $branch + '.remote'))) -AllowedExitCodes @(0, 1)).Text.Trim()
    $reference = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('config', '--get-all', ('branch.' + $branch + '.merge'))) -AllowedExitCodes @(0, 1)).Text.Trim()
    if ([string]::IsNullOrWhiteSpace($remote) -or $remote -eq '.' -or $reference -notmatch '^refs/heads/[^\s]+$') {
        Stop-CalendarUpdate -Code 'UPSTREAM_INVALID' -Message 'The current branch needs one approved remote branch as its upstream.'
    }
    $pushRemote = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('config', '--get', ('branch.' + $branch + '.pushRemote'))) -AllowedExitCodes @(0, 1)).Text.Trim()
    if (-not $pushRemote) {
        $pushRemote = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('config', '--get', 'remote.pushDefault')) -AllowedExitCodes @(0, 1)).Text.Trim()
    }
    if ($pushRemote -and $pushRemote -cne $remote) {
        Stop-CalendarUpdate -Code 'PUSH_DESTINATION_CONFLICT' -Message 'The configured push remote differs from the upstream. Review the intended calendar destination before scheduling this branch.'
    }
    $pushSpecs = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('config', '--get-all', ('remote.' + $remote + '.push'))) -AllowedExitCodes @(0, 1)).Text.Trim()
    if ($pushSpecs -and $pushSpecs -cnotin @(($branch + ':' + $reference), ('refs/heads/' + $branch + ':' + $reference), ('HEAD:' + $reference))) {
        Stop-CalendarUpdate -Code 'PUSH_REFSPEC_CONFLICT' -Message 'The remote has a different or multiple configured push refspecs. Review the intended single calendar branch before scheduling.'
    }
    $urls = @((Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('remote', 'get-url', '--push', '--all', '--', $remote))).Text.TrimEnd([char[]]"`r`n").Split("`n"))
    if ($urls.Count -ne 1 -or [string]::IsNullOrWhiteSpace($urls[0])) {
        Stop-CalendarUpdate -Code 'PUSH_URL_AMBIGUOUS' -Message 'The upstream remote must have exactly one approved push URL.'
    }
    $mirror = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('config', '--bool', '--get', ('remote.' + $remote + '.mirror'))) -AllowedExitCodes @(0, 1)).Text.Trim()
    if ($mirror -eq 'true') { Stop-CalendarUpdate -Code 'MIRROR_UNSUPPORTED' -Message 'A mirror remote cannot be used for a single-calendar update.' }
    return [pscustomobject]@{ Branch = $branch; Remote = $remote; Reference = $reference; Url = $urls[0].TrimEnd("`r") }
}

function Get-CalendarRemoteTip {
    param([string] $GitPath, [string] $Root, [object] $Target)
    $reply = (Invoke-CalendarGit -GitPath $GitPath -Arguments @('-C', $Root, '-c', ('remote.' + $Target.Remote + '.url='), '-c', ('remote.' + $Target.Remote + '.url=' + $Target.Url), 'ls-remote', '--exit-code', '--refs', '--', $Target.Remote, $Target.Reference)).Text.Trim()
    $parts = $reply.Split("`t")
    if ($parts.Count -ne 2 -or $parts[0] -notmatch '^(?:[a-f0-9]{40}|[a-f0-9]{64})$' -or $parts[1] -cne $Target.Reference) {
        Stop-CalendarUpdate -Code 'REMOTE_TIP_INVALID' -Message 'The approved remote did not return exactly one branch tip.'
    }
    return $parts[0]
}

function Assert-CalendarHistory {
    param([string] $GitPath, [object] $Repository, [string] $RemoteTip, [string] $Head)
    $prefix = @('--literal-pathspecs', '-C', $Repository.Root)
    $ancestry = Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('merge-base', '--is-ancestor', $RemoteTip, $Head)) -AllowedExitCodes @(0, 1, 128)
    if ($ancestry.ExitCode -ne 0) {
        Stop-CalendarUpdate -Code 'REMOTE_HISTORY_UNAVAILABLE' -Message 'The remote tip is missing locally, ahead or divergent. Review and synchronize the branch manually before updating.'
    }
    $commits = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('rev-list', ($RemoteTip + '..' + $Head)))).Text.Trim()
    if (-not $commits) { return }
    foreach ($commit in $commits.Split("`n")) {
        $parents = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('rev-list', '--parents', '-n', '1', $commit.Trim()))).Text.Trim().Split(' ')
        if ($parents.Count -ne 2) { Stop-CalendarUpdate -Code 'PENDING_MERGE' -Message 'Unpublished merge/root history requires manual review before calendar publishing.' }
        $names = (Invoke-CalendarGit -GitPath $GitPath -Arguments ($prefix + @('diff-tree', '--no-commit-id', '--no-renames', '--name-only', '-r', '-z', $commit.Trim()))).Text
        foreach ($name in $names.Split([char]0)) {
            if ($name -and $name -cne $Repository.RelativePath) {
                Stop-CalendarUpdate -Code 'PENDING_OTHER_FILES' -Message 'Unpublished commits contain other files. Publish or resolve them manually before calendar updating.'
            }
        }
    }
}

function Assert-CalendarRepositoryReady {
    param([string] $GitPath, [string] $Root)
    foreach ($name in @('MERGE_HEAD', 'CHERRY_PICK_HEAD', 'REVERT_HEAD', 'rebase-merge', 'rebase-apply')) {
        $path = (Invoke-CalendarGit -GitPath $GitPath -Arguments @('-C', $Root, 'rev-parse', '--git-path', $name)).Text.TrimEnd([char[]]"`r`n")
        if (Test-Path -LiteralPath ([IO.Path]::Combine($Root, $path))) {
            Stop-CalendarUpdate -Code 'GIT_OPERATION_PENDING' -Message 'Finish or resolve the current Git merge, rebase, cherry-pick or revert before updating the calendar.'
        }
    }
    $unmerged = (Invoke-CalendarGit -GitPath $GitPath -Arguments @('-C', $Root, 'ls-files', '--unmerged', '-z')).Text
    if ($unmerged.Length -gt 0) { Stop-CalendarUpdate -Code 'INDEX_UNMERGED' -Message 'Resolve unmerged Git index entries before updating the calendar.' }
}

try {
    if (-not $PSCmdlet.ShouldProcess('The configured calendar and its Git repository', 'Export, commit the calendar if changed, and push')) {
        Write-Information -InformationAction Continue -MessageData 'Calendar update skipped. No export, commit or push was attempted.'
        return
    }
    $localState = [Environment]::GetFolderPath('LocalApplicationData')
    if ([string]::IsNullOrWhiteSpace($localState)) { throw 'A writable current-user LocalApplicationData directory is required.' }
    $stateRoot = Join-Path -Path $localState -ChildPath 'osp/consulthelp/update-calendar'
    if ([string]::IsNullOrWhiteSpace($LogDirectory)) { $LogDirectory = Join-Path -Path $stateRoot -ChildPath 'logs' }
    $logPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogDirectory)
    Initialize-CalendarRecord -Directory $logPath
    Write-CalendarRecord -Stage 'dependencies'
    foreach ($name in @('GIT_DIR', 'GIT_WORK_TREE', 'GIT_INDEX_FILE', 'GIT_COMMON_DIR', 'GIT_OBJECT_DIRECTORY', 'GIT_ALTERNATE_OBJECT_DIRECTORIES', 'GIT_NAMESPACE')) {
        if (-not [string]::IsNullOrEmpty([Environment]::GetEnvironmentVariable($name))) {
            Stop-CalendarUpdate -Code 'GIT_CONTEXT_OVERRIDE' -Message 'Use a normal Git worktree context without alternate Git directory, index, object or namespace environment overrides.'
        }
    }
    $exporter = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExporterPath)
    if (-not (Test-Path -LiteralPath $exporter -PathType Leaf)) {
        Stop-CalendarUpdate -Code 'EXPORTER_MISSING' -Message 'outcalget.ps1 was not found. Place it beside this script or supply -ExporterPath.'
    }
    $gitPath = (Get-Command -Name git -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $hostPath = (Get-Command -Name powershell.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-OutputFormat', 'XML', '-File', $exporter, '-Json')
    if ($PSBoundParameters.ContainsKey('Config')) {
        $configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
        $arguments += @('-Config', $configPath)
    }
    if ($PSBoundParameters.ContainsKey('ExportProfile')) {
        $arguments += @('-Profile', $ExportProfile)
    }
    if ($Scheduled) { $arguments += '-Scheduled' }
    if ($script:CalendarDebug) { $arguments += '-Debug' }
    [void][IO.Directory]::CreateDirectory($stateRoot)
    Write-CalendarRecord -Stage 'update-lock'
    try {
        $updateLock = [IO.File]::Open((Join-Path -Path $stateRoot -ChildPath 'update.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    } catch {
        Stop-CalendarUpdate -Code 'UPDATE_LOCK_UNAVAILABLE' -Message 'Another calendar update may be running, or its lock directory is inaccessible. No export or Git work was started.'
    }
    $script:CalendarPendingPath = Join-Path -Path $stateRoot -ChildPath 'git-recovery.pending'
    if (Test-Path -LiteralPath $script:CalendarPendingPath) {
        $script:CalendarRecord.recoveryRequired = $true
        Stop-CalendarUpdate -Code 'GIT_RECOVERY_REQUIRED' -Message 'An interrupted Git command needs review. Check the update record and native helpers before removing git-recovery.pending; no export was started.'
    }
    Write-CalendarRecord -Stage 'export'
    Write-Verbose -Message ('Exporter: ' + $exporter)
    Write-Information -InformationAction Continue -MessageData 'Starting Outlook calendar export. Waiting for outcalget to finish...'
    $output = & $hostPath @arguments
    $exportCode = $LASTEXITCODE
    $script:CalendarRecord.exporterExitCode = $exportCode
    $result = Read-CalendarExportResult -Text ($output -join "`n") -ExitCode $exportCode
    $calendar = Get-PublishedCalendarPath -Result $result
    Write-Information -InformationAction Continue -MessageData ('Outlook calendar export completed: ' + $calendar)
    Write-CalendarRecord -Stage 'git-discovery'
    $repository = Get-CalendarRepository -GitPath $gitPath -CalendarPath $calendar
    Write-CalendarRecord -Stage 'git-preflight'
    Assert-CalendarRepositoryReady -GitPath $gitPath -Root $repository.Root
    $target = Get-CalendarPushTarget -GitPath $gitPath -Root $repository.Root
    $remoteTip = Get-CalendarRemoteTip -GitPath $gitPath -Root $repository.Root -Target $target
    $head = (Invoke-CalendarGit -GitPath $gitPath -Arguments @('-C', $repository.Root, 'rev-parse', 'HEAD')).Text.Trim()
    Assert-CalendarHistory -GitPath $gitPath -Repository $repository -RemoteTip $remoteTip -Head $head
    $gitArguments = @('--literal-pathspecs', '-C', $repository.Root)
    Write-CalendarRecord -Stage 'git-add'
    [void](Invoke-CalendarGit -GitPath $gitPath -Arguments ($gitArguments + @('add', '--', $repository.RelativePath)))
    $diffCode = (Invoke-CalendarGit -GitPath $gitPath -Arguments ($gitArguments + @('diff', '--cached', '--quiet', '--', $repository.RelativePath)) -AllowedExitCodes @(0, 1)).ExitCode
    $script:CalendarRecord.calendarChanged = $diffCode -eq 1
    if ($diffCode -eq 1) {
        $script:CalendarRecord.gitOutcomeUncertain = $true
        Write-CalendarRecord -Stage 'git-commit'
        [void](Invoke-CalendarGit -GitPath $gitPath -Arguments ($gitArguments + @('commit', '--only', '-m', 'Update calendar', '--', $repository.RelativePath)))
        $script:CalendarRecord.commitCreated = $true
    }
    $head = (Invoke-CalendarGit -GitPath $gitPath -Arguments @('-C', $repository.Root, 'rev-parse', 'HEAD')).Text.Trim()
    $script:CalendarRecord.commit = $head
    $script:CalendarRecord.gitOutcomeUncertain = $false
    if ($head -cne $remoteTip) {
        Write-CalendarRecord -Stage 'git-push-check'
        $currentTarget = Get-CalendarPushTarget -GitPath $gitPath -Root $repository.Root
        foreach ($key in @('Branch', 'Remote', 'Reference', 'Url')) {
            if ($currentTarget.$key -cne $target.$key) { Stop-CalendarUpdate -Code 'PUSH_TARGET_CHANGED' -Message 'The Git branch or push destination changed during the update. Nothing was pushed.' }
        }
        Assert-CalendarHistory -GitPath $gitPath -Repository $repository -RemoteTip $remoteTip -Head $head
        $script:CalendarRecord.gitOutcomeUncertain = $true
        Write-CalendarRecord -Stage 'git-push'
        [void](Invoke-CalendarGit -GitPath $gitPath -Arguments @('-C', $repository.Root, 'push', '--porcelain', '--no-follow-tags', '--recurse-submodules=no', '--', $target.Remote, ($head + ':' + $target.Reference)))
        Write-CalendarRecord -Stage 'git-push-verify'
        if ((Get-CalendarRemoteTip -GitPath $gitPath -Root $repository.Root -Target $target) -cne $head) {
            Stop-CalendarUpdate -Code 'PUSH_NOT_CONFIRMED' -Message 'The remote tip could not be confirmed as the calendar commit. Inspect the remote before retrying.'
        }
        $script:CalendarRecord.pushed = $true
        $script:CalendarRecord.remoteConfirmed = $true
        $script:CalendarRecord.gitOutcomeUncertain = $false
        Write-Output -InputObject 'Calendar updated, committed and pushed.'
    } else {
        $script:CalendarRecord.remoteConfirmed = $true
        Write-Output -InputObject 'Calendar unchanged. Remote already contains the current commit.'
    }
    if ($script:CalendarRecord.recoveryRequired) {
        Stop-CalendarUpdate -Code 'GIT_RECOVERY_REQUIRED' -Message 'The remote was checked, but the Git recovery marker could not be cleared. Review the marker before the next update.'
    }
    $script:CalendarRecord.outcome = 'success'
    $updateHandled = $true
} catch {
    $updateExit = 1
    $updateHandled = $true
    $message = 'Calendar update failed before export. Check dependencies and supplied paths.'
    if ($null -ne $script:CalendarRecord) {
        $script:CalendarRecord.outcome = 'failed'
        $script:CalendarRecord.errorCode = $script:CalendarRecord.stage.ToUpperInvariant().Replace('-', '_') + '_FAILED'
        $message = 'Calendar update failed during ' + $script:CalendarRecord.stage + '. Review its run record; the export or a local commit may already exist.'
        if ($_.Exception.Data.Contains('CalendarCode')) {
            $script:CalendarRecord.errorCode = [string]$_.Exception.Data['CalendarCode']
            $message = $_.Exception.Message
        }
        if ($script:CalendarDebug) { $script:CalendarRecord.diagnostics.exceptionType = $_.Exception.GetType().FullName }
    }
    Write-Error -Message $message -ErrorAction Continue
} finally {
    if ($null -ne $script:CalendarLog) {
        try {
            $script:CalendarRecord.completed = $updateHandled
            if ($updateHandled) { $script:CalendarRecord.exitCode = $updateExit }
            Write-CalendarRecord
        } catch {
            $updateExit = 1
            Write-Warning -Message 'The calendar update completion record could not be saved. The business outcome may already be committed; inspect Git and the exporter record.'
        } finally {
            $script:CalendarLog.Dispose()
        }
    }
    if ($null -ne $updateLock) { $updateLock.Dispose() }
}
exit $updateExit
