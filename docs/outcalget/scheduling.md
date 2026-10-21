# Schedule and monitor update-calendar

[Documentation index](README.md) · [Wrapper behavior, parameters and recovery](update-calendar.md)

Schedule `update-calendar.ps1` to export, commit the calendar when changed, and push its branch. Keep `outcalget.ps1` beside it (or use ExporterPath). Both may live in `scripts/` or an independent tools folder; documentation/XML files are optional. These steps do not require administrator rights, but your employer must permit task creation and the existing script/Outlook/Git controls.

## 1. Prepare the signed-in user session

1. Use Windows PowerShell 5.1 and Outlook Classic under your ordinary Windows account. Complete Outlook profile selection, sign-in and security dialogs interactively first. The task will run only while this user is logged on; logged-off/SYSTEM/S4U automation is unsupported. Test a locked desktop separately if needed.
2. Keep the existing working exporter configuration. For a new configuration only, run `outcalget.ps1 -Configure`, optionally with `-Profile team`, interactively. Configure the output inside the intended Git worktree and verify the actual source/content/date range. Do not put Configure in the scheduled action.
3. Confirm Git is on this user's PATH in a fresh NoProfile PowerShell process, the repository already has the correct identity/upstream, and its signing/hooks/approved credentials work without prompts. Review locally with `git -C "C:\repo" status` and `git -C "C:\repo" branch -vv`; do not share private output. The wrapper pushes only the current branch to its upstream branch at that remote's single push URL. It refuses unrelated unpublished commits, unfinished Git operations, mirror remotes, conflicting push-remote/refspec overrides and ambiguous configuration. It does not automatically fetch or repair a branch.
4. Run the exact scheduled command once from a normal user session. This performs a real export, commit if changed and push; use only your intended calendar and approved destination. Read its wrapper record and inspect the published calendar in its intended consumer before relying on automation.

Example (substitute your paths and omit Profile for the existing default calendar):

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\repo\scripts\update-calendar.ps1" -Scheduled -Profile team -Debug
```

Debug adds diagnostics, not a dry run. `-WhatIf` skips the operation and cannot establish Outlook, credentials or remote readiness.

## 2. Create the task in Windows Task Scheduler

Use **Create Task** with an identifiable name, for example `Calendar update`.

**General**

- Select the same ordinary Windows user who configured Outlook and the export.
- Select **Run only when user is logged on**.
- Leave **Run with highest privileges** unchecked. Do not use SYSTEM or a stored-password/S4U/logged-off mode.

**Triggers**

- Choose your desired schedule, for example daily at 09:00, optionally repeating every hour.
- If using an at-logon trigger, delay it by two minutes so the desktop and credentials can initialize.
- Stagger separate calendar tasks: the wrapper allows one update at a time per user, even with different profiles.

**Actions → Start a program** (three separate fields):

```text
Program/script:
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

Add arguments:
-NoLogo -NoProfile -NonInteractive -STA -File "C:\repo\scripts\update-calendar.ps1" -Scheduled -Profile team -Debug

Start in:
C:\repo
```

Use your approved interpreter and absolute script paths. Start in is the directory, without surrounding quotes. Remove `-Profile team` for the original default configuration, or replace it with `-Config "C:\PrivateState\config.json"`; do not combine selectors. Append `-ExporterPath "C:\Tools\outcalget.ps1"` only for a nonadjacent exporter. Optional `-LogDirectory "C:\PrivateState\CalendarUpdateLogs"` changes wrapper logs only; use a private writable directory outside the repository. For multiple tasks, distinct wrapper log directories make correlation easier. Add WindowStyle Hidden before File only after acceptance if you want to hide the PowerShell window; Office dialogs cannot be suppressed this way.

Do not append shell redirects (`>`/`2>`), `-Command`, execution-policy bypass or `-Confirm:$false` to this Exec action. The script owns its run logs. No extra launcher or batch file is needed.

**Conditions / Settings**

- If already running: **Do not start a new instance**.
- Leave automatic failure retries disabled. Review an error or uncertain native operation before retrying.
- Enable **Run task as soon as possible after a scheduled start is missed** if desired; it still needs the signed-in online session. This cannot export while logged off/asleep.
- Avoid idle-only conditions. Allow battery execution and disable stops on battery transitions if your usage/policy permits. Do not wake the laptop by default.
- A network condition is optional; the exporter checks Windows Internet status itself. Git authentication/reachability can still fail independently.
- Use at least **two hours** for an external runtime stop limit at default export/Git budgets and a short pending calendar history; increase it if your approved native helpers/budgets need more. Git's default wait is 120 seconds per process, not a total job deadline. A Scheduler stop can interrupt cleanup: missing completion evidence is UNKNOWN, and retained markers require review.

Save the task. No script automatically registers or modifies tasks, Event Log sources, security settings or credentials.

## 3. Test the saved action

Right-click the task → **Run**. Wait until it stops showing Running, refresh, and check Last Run Time and Last Run Result. `0x0` is necessary but must agree with the current script record. Repeat once using a short temporary trigger while signed in, then restore your desired trigger. Existing task History can help with launches if your organization enables it; no administrator-only logging setup is required by the scripts.

On the target laptop test initially open Outlook (left open), initially closed Outlook (owned startup/sync/normal close), a locked desktop if required, and Internet/authentication failure. Failure must retain the prior good output or disclose a calendar already published before a later failure. Confirm event content in the intended consumer; native sync completion is not independent proof of shared-calendar server freshness. Do not clear Outlook recovery markers to make a test pass; follow [exporter recovery](operations.md#recovery).

## 4. Monitor the complete update

Task status (from the same user; substitute the task folder/name if different):

```powershell
Get-ScheduledTaskInfo -TaskName 'Calendar update' |
    Select-Object LastRunTime, LastTaskResult, NextRunTime
```

Latest wrapper record (use your LogDirectory override instead if configured):

```powershell
$logDirectory = Join-Path $env:LOCALAPPDATA 'osp\consulthelp\update-calendar\logs'
$file = Get-ChildItem -LiteralPath $logDirectory -Filter 'update-calendar-*.json' -File |
    Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
if ($null -eq $file) { throw 'No update record found. Inspect the task action and host startup.' }
$run = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
$run | Select-Object startedUtc, updatedUtc, completed, outcome, exitCode, stage,
    exporterRunId, exporterExitCode, exportPublished, calendarChanged,
    commitCreated, commit, pushed, remoteConfirmed, gitExitCode,
    gitOutcomeUncertain, recoveryRequired, errorCode
```

Check that startedUtc belongs to the latest task launch, rather than an older successful run. A record is rewritten during checkpoints; if read while being rewritten, retry the read. Unreadable or incomplete data cannot establish success. For multiple tasks, use each task's own log directory or correlate its launch time; profile names are deliberately absent from logs.

A confirmed completed update has `completed=true`, `outcome=success`, `exitCode=0`, `exportPublished=true`, `remoteConfirmed=true`, `recoveryRequired=false`, with Scheduler result `0x0`. `pushed=true` means a push was performed and verified. `pushed=false` with those success fields means the remote already held HEAD after a fresh unchanged export. The commit hash identifies the revision, including a retried existing commit.

To locate the corresponding exporter record when an exporterRunId is present:

```powershell
$exportLogs = Join-Path $env:LOCALAPPDATA 'osp\consulthelp\outcalget\logs'
Get-ChildItem -LiteralPath $exportLogs -Filter ("outcalget-*-{0}.json" -f $run.exporterRunId) -File
```

If the exporter uses a separately configured log directory, look there. Inspect its completion/publication flags, synchronization outcome, warnings and stable error codes. Exporter success alone does not confirm the commit/push. Keep configuration, mailbox identities and ICS content private; review any log before sharing it.

## 5. Respond to failures

- **No new wrapper record:** verify absolute interpreter/script paths, user session, signing/execution/application policy, task launch result and the private log directory's permissions. Parse/binding/policy/host-start failures can happen before logging.
- **Export failure:** use exporterRunId and its record. There is no Git publication for a failed/malformed export. A later export cleanup failure can still have published the output, so inspect the flag and recovery state before retrying.
- **Git preflight failure:** errorCode identifies unsupported context, pending operation/history or remote configuration. Resolve through your normal approved repository maintenance process. The wrapper never fetches, merges, resets or force-pushes.
- **Ordinary commit/signing/hook/push refusal:** inspect Git locally using approved commands and credentials. The published file, index change or calendar commit may remain. A normal rejected push preserves the commit; once its cause is resolved, rerun the task. An unchanged re-export still pushes pending calendar-only history.
- **Timeout, interrupted Git work or recoveryRequired:** disable the task, inspect the record and ensure the owned Git command and any hook/credential/signing children have finished. Check actual local/remote history and native Git locks. Then remove only the reviewed `%LOCALAPPDATA%\osp\consulthelp\update-calendar\git-recovery.pending` marker and re-enable/run the task. The marker contains the wrapper run ID. Do not remove a live Git index lock or the exporter's recovery artifacts. The harmless `update.lock` file need not be deleted.
- **Stale/incomplete record, abrupt stop or failed final log:** outcome is UNKNOWN even if Scheduler shows zero or the ICS exists. Inspect the remote, local commit and exporter evidence before retrying. No automatic publication/cleanup success is assumed.

Development verification used synthetic calendars and filesystem-only remotes on Linux. Actual Windows PowerShell 5.1, Outlook/Exchange, Task Scheduler, credentials/signing and endpoint acceptance must be established on the intended Windows laptop using the steps above.
