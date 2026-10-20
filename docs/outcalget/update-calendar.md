# Calendar Git update helper

[Documentation index](README.md)

Implementation: [`update-calendar.ps1`](../../outcalget/update-calendar.ps1). Optional [native help](update-calendar-help.xml).

## Location and prerequisites

The wrapper can live in a repository root, a `scripts/` subdirectory, or an independent tools directory, and can be invoked from any current directory. It never changes the caller’s PowerShell location. By default it requires `outcalget.ps1` beside it; `-ExporterPath` selects a nonadjacent exporter. Neither script requires Markdown, XML help or agent instructions to execute.

Requires Windows PowerShell 5.1 (`powershell.exe`), installed Git, the exporter’s declared Windows/Outlook dependencies and an existing export configuration. The configured output must be inside a Git working tree whose current branch has an approved upstream/push destination and authentication. The wrapper does not configure Outlook, create a repository, choose a new remote or provision credentials.

The exporter’s successful `-Json` result selects the actual absolute output file. There is no assumption that it is named `calendar.ics`, beside the scripts, or in the current directory. Existing default configuration is used if neither `-Config` nor `-Profile` is supplied. `-Profile` names a saved export configuration, not an Outlook MAPI profile. Missing/invalid configurations fail through the exporter without fallback or setup prompting.

```powershell
# These commands can be invoked from an unrelated directory.
powershell.exe -NoLogo -NoProfile -File "C:\repo\scripts\update-calendar.ps1"
powershell.exe -NoLogo -NoProfile -File "C:\repo\scripts\update-calendar.ps1" -Profile team
powershell.exe -NoLogo -NoProfile -File "C:\Tools\update-calendar.ps1" -ExporterPath "C:\OtherTools\outcalget.ps1" -Config "C:\PrivateState\config.json"
```

`-ExporterPath` and `-Config` accept absolute paths or paths relative to the caller’s current PowerShell filesystem location. The wrapper makes a relative configuration path absolute before invoking the exporter. `-Config` and `-Profile` are mutually exclusive. The exporter’s saved output and account-bound configuration do not move when the scripts move.

## Export, commit and failure behavior

The wrapper announces the Outlook export before launching the child and reports the published output path after it validates success. It invokes the exporter with `-STA`, `-NonInteractive` and `-Json`, checks both process status and the published-result envelope, then finds the Git worktree containing that output. The JSON mode suppresses the child’s individual phase messages; a quiet interval after the start message means the wrapper is waiting for the exporter. `-Verbose` shows the resolved exporter path. It uses `git -C` and a literal single-file pathspec to stage and commit only that calendar. An unchanged calendar produces no commit or push. Unrelated staged files remain staged. Native hooks, signing and the repository’s configured `git push` behavior remain active.

The previous calendar is never pre-deleted. The exporter retains its own atomic-publication/recovery protections. If export fails, the wrapper performs no staging, commit or push; inspect the exporter’s run record if it reports publication before a later failure. A missing/invalid successful result also prevents Git work. Handled exporter failures display the exporter’s error code/message and run-record path when provided; warnings are surfaced. If the exporter reports publication before a later failure, the wrapper calls that out and stops before Git work. Malformed/raw child output is not echoed. If output is outside any Git repository, the export may already have completed, but no Git publication follows. Git failures can leave the calendar staged or a commit local; a push failure explicitly reports that the commit remains local. Errors exit 1; success/preview exits 0.

`-WhatIf` or declined `-Confirm` skips export and Git mutations and prints an explicit skipped message. Dependency/path resolution can still run; preview does not establish Outlook access or remote readiness.

```powershell
powershell.exe -NoLogo -NoProfile -File "C:\repo\scripts\update-calendar.ps1" -WhatIf
```

## Implementation and verification, 19 October 2026

`Get-PublishedCalendarPath` validates the successful exporter envelope and existing absolute output path. The child host’s XML transport preserves the JSON string, including Unicode paths, across the native PowerShell boundary. `Get-CalendarRepository` uses Git’s ASCII parent-directory result (`rev-parse --show-cdup`) with the already-known calendar directory; it avoids interpreting a Unicode root path through a legacy console code page. All Git file operations use `--literal-pathspecs`. Public ShouldProcess covers the complete export/commit/push operation.

**PASS:** synthetic calendar fixture families for scripts at the repository root, under `scripts/`, and outside the repository, invoked from an unrelated PowerShell location and process directory. These exercise actual local Git staging, commit and push against filesystem-only bare remotes, Unicode/spaces/brackets in the output path, unchanged output, preservation of unrelated staged files and caller location, failed export, malformed/missing/unpublished results, WhatIf, a nonadjacent exporter and relative explicit config/profile forwarding. Additional fixtures verify commit-hook refusal, rejected push with the commit retained locally, a calendar directly at the repository root, and refusal of a calendar outside a repository. Native help works with and without XML. Parser/default plus Windows PowerShell 5.1 syntax/command/type analysis has zero findings.

The fixture replaces only the Outlook/export process boundary with a synthetic exporter and a local `powershell.exe` adapter that delegates to PowerShell 7.6.6 while omitting Linux’s unsupported `-STA`. No production test switch or dependency was added. `GIT_ALLOW_PROTOCOL=file` restricted Git transport. No real calendar or external destination was used. **NOT RUN:** Windows PowerShell 5.1/Outlook/Git-for-Windows, real signing/credentials/remote protection, Task Scheduler and endpoint acceptance. These Linux fixtures establish the tested path and Git behavior, not Windows business acceptance.

## Troubleshooting an apparently skipped export

With both scripts in `scripts/`, invoke `.\scripts\update-calendar.ps1` from the repository root. The adjacent exporter is selected by `$PSScriptRoot`; the current directory does not select a different exporter. The default saved configuration is used unless you specify `-Profile` or `-Config`.

Expected wrapper messages are “Starting Outlook calendar export”, then “Outlook calendar export completed” and the actual output path, followed by the Git outcome. “Calendar unchanged” means the successful exporter result was validated and Git found no staged content difference. It does not mean the export was skipped. An already-open Outlook session is reused and may show no visible UI activity. “Calendar update skipped” means ShouldProcess declined or previewed the operation; no child export was launched.

To see the exporter’s own phases and error directly, run an export without the Git wrapper:

```powershell
powershell.exe -NoLogo -NoProfile -STA -File .\scripts\outcalget.ps1
# If you normally use a named configuration, append -Profile team.
```

A successful run reports the actual output path and run-record location. Review that output locally; share only error codes/messages, with private paths omitted. Avoid treating the absence of a new Outlook window as proof that no export occurred.
