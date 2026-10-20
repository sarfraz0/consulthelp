# sendshards operations

[Documentation index](README.md)

## Current audit status

Version **1.0.2** implements the lost-checkpoint and exited-parent cleanup findings from the [20 October audit](audit-2026-10-20.md). Passing local regressions do not establish Windows/endpoint acceptance; that release gate remains pending.

Version 1.0.1 fixes the six findings in the [18 October audit](audit-2026-10-18.md): initial planning, checkpoint replacement, escaped JSON keys, manifest types, no-work index checks and timestamp round trips. See [current verification](verification.md#fixes-and-relocation-verification-19-october-2026) for synthetic execution evidence and remaining Windows gaps. After any checkpoint failure, still inspect actual Git/remote state: a successful push can precede a failed state write.

## Extracted operational notes

WhatIf/declined confirmation: read-only local checks, no network, logs, locks,
directories, staging, or checkpoint writes. Debug never introduces prompting.
Exit: 0 complete/requested preview; 1 unexpected; 2 invalid input/plan;
3 operational/dependency/repository conflict; 4 controlled timeout.
Binding, execution-policy refusal, Ctrl+C, and hard termination can occur outside
this mapping. Missing completion is an unknown outcome; resume reconciles Git.
Audit initialization must succeed before mutations. Later logging failure warns. Cleanup uncertainty preserves a primary failure; after an otherwise successful Git call, unconfirmed cleanup stops transfer with OwnedCleanup. Checkpoint failure stops further work.
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

## Plan, resume and recovery boundaries

Read the behavior contract in the [documentation index](README.md) before using either mode. Plan may create its own JSON/log/lock state after confirmation, but performs no network operation and never stages, commits or pushes shards. Push is an explicit publication workflow against the configured approved remote, requiring Windows, Git for Windows 2.30 or newer, an existing remote branch and pre-provisioned credentials.

Resume uses the same plan and original persisted time budget. Inspect native Git locks and repository/checkpoint state after an interruption; do not delete locks while Git may still run, restart the original budget or blindly repeat an uncertain push. A failed commit can leave the next shard staged, and a successful commit can remain local until a verified push. Other Git applications are not locked out by this tool's own common-directory lock.

Embelish emits no checksum manifest in its current revision. Sendshards accepts those legacy sets with a warning; it does not decrypt or prove archive authenticity. Exact Git blob verification checks the transferred bytes and rejects byte-changing filters, including LFS pointers. Hooks, signing, credential helpers, TLS, proxy and endpoint controls remain active.

The script location and caller’s directory can differ from the source repository. Source/plan arguments select the data, and persisted plans retain canonical absolute paths. Moving the script does not move source files or make an edited active plan valid. Native help XML remains optional. Windows path, process-tree, signing/credential and scheduler acceptance still require target execution.

## Original start record and missing state

Keep `<PlanPath>`, `<PlanPath>.state.json` and `<PlanPath>.start.json` together and back them up securely. The new start record contains schema version, plan identity/hash and the exact original StartedUtc; it is created once before remote contact or Git mutations, using the same bounded atomic JSON writer. It is never rewritten during resume. A valid older checkpoint without that record is upgraded using its existing start, with no new hour. Plan refuses surviving state/start filenames.

If state is missing but a start record or exact local transfer commits remain, Push stops with MissingCheckpoint. If the start and checkpoint disagree, it stops with ChangedStart. Restore the matching original checkpoint/start record and inspect local/remote tips before resuming. Corrupt records fail validation; they are never treated as absent. Do not delete runtime data or edit timestamps to restart an expired plan. If all original runtime data is lost and there is no history proving prior activity, the original start cannot be reconstructed; custody/backups are required. Plan creation time is not substituted for the first Push time.

## Native Git lifetime

Windows 10/11 Job Objects replace taskkill/PID-only cleanup. The in-file adapter creates a kill-on-close job and supplies both an explicit inherited-handle list and job list to CreateProcess. Git starts suspended inside that job, then resumes only after pipe ownership is established. This avoids a helper-spawn race and the suspended-orphan window during abrupt host termination. Ordinary child processes inherit job membership; parent exit does not release them. No breakaway flag, process-name kill, uncontained fallback, elevation or endpoint-policy change is used.

Every invocation terminates and closes its own job in finally, checking active-process accounting for up to five seconds before releasing handles. SafeHandle ownership is also a finalization/host-exit backstop. Captured stdout/stderr remain independently bounded to 2 MiB. Native hooks, signing, environment, approved credentials and TLS/proxy controls stay active; stdin receives EOF and GIT_TERMINAL_PROMPT remains 0. Helpers must finish their work within the owning Git invocation. Processes deliberately launched through an external broker/WMI or an existing service are outside the ordinary job descendant contract; use approved helpers and validate their actual behavior.

Inline native compilation, job creation/nesting or file APIs blocked by application control cause a refusal; there is no alternate launch path. Windows PowerShell 5.1, Git-for-Windows helpers/signing/hooks, parent-exit/inherited-pipe cleanup, Ctrl+C and hard host termination must be exercised on the managed target before production release. Linux protocol regressions use a development process-boundary adapter and do not test the Windows job implementation.
