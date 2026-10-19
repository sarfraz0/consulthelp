# sendshards operations

[Documentation index](README.md)

## Current audit status

Version 1.0.1 fixes the six findings in the [18 October audit](audit-2026-10-18.md): initial planning, checkpoint replacement, escaped JSON keys, manifest types, no-work index checks and timestamp round trips. See [current verification](verification.md#fixes-and-relocation-verification-19-october-2026) for synthetic execution evidence and remaining Windows gaps. After any checkpoint failure, still inspect actual Git/remote state: a successful push can precede a failed state write.

## Extracted operational notes

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

## Plan, resume and recovery boundaries

Read the behavior contract in the [documentation index](README.md) before using either mode. Plan may create its own JSON/log/lock state after confirmation, but performs no network operation and never stages, commits or pushes shards. Push is an explicit publication workflow against the configured approved remote, requiring Windows, Git for Windows 2.30 or newer, an existing remote branch and pre-provisioned credentials.

Resume uses the same plan and original persisted time budget. Inspect native Git locks and repository/checkpoint state after an interruption; do not delete locks while Git may still run, restart the original budget or blindly repeat an uncertain push. A failed commit can leave the next shard staged, and a successful commit can remain local until a verified push. Other Git applications are not locked out by this tool's own common-directory lock.

Embelish emits no checksum manifest in its current revision. Sendshards accepts those legacy sets with a warning; it does not decrypt or prove archive authenticity. Exact Git blob verification checks the transferred bytes and rejects byte-changing filters, including LFS pointers. Hooks, signing, credential helpers, TLS, proxy and endpoint controls remain active.

The script location and caller’s directory can differ from the source repository. Source/plan arguments select the data, and persisted plans retain canonical absolute paths. Moving the script does not move source files or make an edited active plan valid. Native help XML remains optional. Windows path, process-tree, signing/credential and scheduler acceptance still require target execution.
