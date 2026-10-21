# Calendar Git update helper

[Documentation index](README.md) · [Task Scheduler setup and monitoring](scheduling.md)

Implementation: [`update-calendar.ps1`](../../outcalget/update-calendar.ps1). Update run-record version: 1.1.0. Optional [native help](update-calendar-help.xml).

## Deployment and prerequisites

Run the wrapper from any current directory, with the pair at a repository root, under `scripts/`, or in an independent tools directory. It never changes the caller's location. `outcalget.ps1` must be beside it unless selected with `-ExporterPath`. Markdown, XML and agent files are optional runtime documentation.

Requires Windows PowerShell 5.1, Git on the task user's PATH, Outlook Classic and an existing account-bound exporter configuration. Use the ordinary configured Windows user, logged on in the same desktop session. Logged-off/SYSTEM/service automation is unsupported. PowerShell `-NonInteractive` cannot suppress Office sign-in or security dialogs. No elevation, installation, policy bypass or security-control changes are performed.

The exporter returns the actual absolute published calendar path. That file must belong to an existing Git worktree on a branch with one approved upstream remote/branch and one push URL. Existing Git identity, signing, hooks, authentication, proxy and TLS controls must already work without prompts. Use a dedicated calendar branch/worktree where practical. Do not run other Git commands or direct exports to the same output during an update.

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\repo\scripts\update-calendar.ps1" -Scheduled
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\repo\scripts\update-calendar.ps1" -Scheduled -Profile team -Debug
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\Tools\update-calendar.ps1" -ExporterPath "C:\OtherTools\outcalget.ps1" -Config "C:\PrivateState\config.json"
```

## Parameters

| Parameter | Meaning |
|---|---|
| `-ExporterPath` | Defaults to adjacent `outcalget.ps1`; relative paths resolve from the caller's filesystem location. |
| `-Config` | Explicit exporter configuration, resolved from the caller's filesystem location and passed as an absolute path. Mutually exclusive with Profile. |
| `-Profile` / `-ExportProfile` | Saved export configuration name, not an Outlook MAPI profile. Omit both selectors for the existing default configuration. |
| `-Scheduled` | Marks the wrapper record and forwards Scheduled to the exporter; does not register a task. |
| `-LogDirectory` | Wrapper records only. Default: `%LOCALAPPDATA%\osp\consulthelp\update-calendar\logs`. Relative paths follow the caller. Does not relocate exporter records/configuration or the update lock/recovery marker. |
| `-GitTimeoutSeconds` | Per Git process wait, default 120, range 15–1800 seconds, plus up to two seconds for redirected stream completion. Does not change export/Outlook budgets. |
| `-Debug` | Non-prompting wrapper diagnostics and exporter Debug; same business operation. |
| `-WhatIf`, `-Confirm` | Native approval for the complete operation; preview or decline skips export, logs, locks and Git. |

Missing named profiles fail through the exporter without a default fallback or interactive setup. Detailed help is optional; `Get-Help -Full` falls back to generated syntax without XML.

## Export, commit and push

The wrapper initializes a required run record and acquires an exclusive per-user update handle before export. All wrapper updates under that user are serialized by refusal, including different tasks/configurations. The persistent `update.lock` file is harmless after its handle is released. Stagger multiple tasks. The exporter's own concurrency, locks and recovery rules remain intact for direct exporter use.

The wrapper announces export, starts the child with NoProfile/NonInteractive/STA and Json, then validates both its process exit and published-result envelope. Scheduled and Debug are forwarded when selected. JSON mode suppresses individual child phases. Warnings, safe handled exporter errors and its run-record location remain visible. Failure or malformed output prevents all Git work; publication followed by a later exporter failure is disclosed. The previous calendar is never pre-deleted.

Before staging, the wrapper rejects unresolved Git operations, unmerged index entries, detached/missing/ambiguous upstream configuration, mirror remotes, conflicting push remote/refspec overrides and ambiguous push URLs. Alternate Git directory/index/object/namespace environment overrides fail before export. The push endpoint is inspected with `ls-remote` as part of this push workflow; no fetch, pull, merge, remote creation or credential provisioning is performed. The remote tip must exist locally and be an ancestor of HEAD. Pending commits may change only this calendar and must have linear history; other unpublished content requires manual repository maintenance.

Only the literal calendar file is staged and committed (`commit --only`); unrelated staged files stay staged. A changed calendar gets an `Update calendar` commit. Unchanged output creates no commit. A pending calendar-only commit is still pushed after an unchanged export, fixing retries following rejected pushes. A synchronized unchanged export succeeds after confirming the remote tip.

Push names the approved upstream remote and explicit commit/branch refspec, with tags and submodule pushes excluded. A conflicting push-remote override or different/multiple explicit push refspecs is refused instead of silently selecting another destination. This intentionally narrows the old bare `git push` behavior; review the existing upstream as the intended task destination. The branch and destination are rechecked before push; hooks, commit signing, remote-specific credentials/proxy settings, TLS validation and server protections remain active. The remote tip is inspected again before recording confirmed success. Separate fetch/push URLs are supported; the read-only query temporarily selects the resolved push URL without modifying saved configuration.

Git stdin is closed and child-only `GIT_TERMINAL_PROMPT=0` / `GCM_INTERACTIVE=Never` prevent Git/GCM credential prompts. Other signing/SSH/hook dialogs may still occur and must be provisioned in advance; they share the Git wait limit. Raw Git stdout/stderr are drained concurrently and never printed or retained in the run record. Controlled, reviewed native helpers are assumed; their redirected output is buffered in memory and is not a hostile-output memory sandbox.

## Outcomes, logs and recovery

Success/preview exits 0; handled failure exits 1. Host parse/binding/signing/policy refusal can precede the script's logging and exit mapping. Export failure does not stage/commit/push. Git failure can leave the new calendar published, staged, or committed locally; the log distinguishes these phases. A rejected push retains the commit. After the remote refusal is resolved, rerunning exports again and pushes pending calendar-only commits even when output is unchanged.

Normal and Debug runs use the same unique UTF-8 JSON record, checkpointed before later mutations and finalized on handled completion. Fields include UTC times, stage, completion/outcome/exit code, exporter run ID and exit/publication flags, calendar-change/commit flags, commit hash, push/remote confirmation, Git exit status, uncertainty/recovery flags and stable error code. Debug adds PowerShell version, script hash when available and failure exception type. No calendar contents, profile names, remote URLs, mailbox identities, paths, raw exceptions or raw native output are logged. Files inherit the chosen directory's permissions; choose an approved user-private log directory. Best-effort retention keeps at most 200 older records plus the current run, and removes closed records older than 14 days; active/inaccessible files are preserved with warnings.

Each Git command has a `git-recovery.pending` marker under `%LOCALAPPDATA%\osp\consulthelp\update-calendar`. Normal native completion clears it, including ordinary refusals. Timeout, interrupted work or unclosed pipes retain it and block subsequent wrapper runs before export. Only the exact started Git process may be stopped on timeout; descendants are not claimed stopped. Disable the task, let/reconcile native Git, hook, credential and signing work finish, inspect the correlated record and repository/remote, then remove only that reviewed marker. A Git-owned index lock or recovery state must be reviewed separately; never delete an active lock. Do not automatically clear markers, reset commits, force push, disable hooks/signing or change TLS/proxy/security controls.

A missing, incomplete, unreadable or stale record means UNKNOWN. Cancellation/hard termination does not guarantee finally/cleanup or a nonzero Windows host status. Completion logging failure returns failure even if export/push already succeeded. Correlate the current wrapper record, Task Scheduler result and exporter record; an old ICS timestamp or successful exporter record alone cannot confirm a full update. See [scheduling and monitoring](scheduling.md) for commands and acceptance checks.

## Verification

The [21 October 2026 verification](verification.md#scheduled-update-audit-21-october-2026) records local synthetic regression and scheduling checks separately from actual Windows/Outlook acceptance. No real calendar, external remote or registered task was used. Exporter version 0.3.5 and its implementation are unchanged.
