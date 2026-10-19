# outcalget instructions

Read the root `AGENTS.md`, `docs/Powershell guidelines.md`, and `docs/outcalget/README.md` before making changes. These instructions refine the shared policy for the exporter and Git update helper.

## Existing contracts

- Preserve Outlook Classic's native `GetCalendarExporter`/`SaveAsICal` route, full details, today through today + 29 local dates inclusive, unrestricted working hours and saved privacy/attachment choices.
- Preserve existing schema-1 configurations and exact schema-2 source bindings. `-Profile` names an export configuration; it does not select an Outlook MAPI profile. A missing named profile must fail without default fallback or an interactive setup prompt.
- Preserve folder/store/shared-owner identity checks, including `NameSpace.CompareEntryIDs` and the validated shared-default null-Store exception. Do not substitute cache scraping or a fabricated exporter.
- A user-open Outlook session remains open and receives no explicit synchronization. Only the exact process started by the exporter may receive normal `Quit`, after readiness, synchronization, usage-lease and user-activity checks. Never force-close Outlook.
- Preserve coordination gates, per-source/config/output locks, output claims, correlated worker replies, bounded budgets, synchronization-setting restoration and recovery markers. Do not delete uncertain staging or recovery artifacts to make a retry succeed.
- Keep native synchronization completion distinct from independently verified server freshness and consumer acceptance.
- Preserve `-Json` envelopes, non-prompting `-Debug`, `-WhatIf`/`ShouldProcess`, existing exit codes, private run records and allowlisted diagnostics. Do not log calendar content, owner addresses, native IDs, profile names or raw exceptions.

## Files and documentation

- Scope normally includes `outcalget.ps1`, `update-calendar.ps1` and `docs/outcalget/`. Shared guidance changes require task scope that includes them.
- `outcalget.ps1` uses `../docs/outcalget/outcalget-help.xml` for native help. Preserve relative help resolution, `Get-Help -Full`, `-?`, `-Help` and `-Version`; keep the XML and Markdown consistent.
- Implementation rationale and existing complexity/length exceptions live in `docs/outcalget/implementation-notes.md`, keyed by function. Historical original line numbers do not describe the current source.
- `update-calendar.ps1` exports, stages, commits and pushes the actual published output returned by outcalget. Never execute it against real calendar data during development; behavior checks must use copied scripts, synthetic exporter results and filesystem-only remotes. Preserve prior output on export failure, caller location, literal single-file staging and native hooks/push controls. Its adjacent exporter dependency can be overridden with `-ExporterPath`; see `docs/outcalget/update-calendar.md`.

## Checks

- For documentation-only changes, compare noncomment/nonwhitespace PowerShell token kind and text against the original, preserve embedded native strings and file encodings, check local documentation links, and verify native help and `-Version` in fresh processes from another directory.
- For behavior changes, use synthetic fixtures to assert affected success and failure paths. Include prior-output preservation, source identity, timeout, overlap, interrupted synchronization and cleanup when relevant.
- Do not dot-source the complete exporter for tests; its entry point performs business work and exits. Load selected AST functions in isolation when necessary.
- Never execute corporate exports, Git publishing or recovery cleanup merely to validate an edit. Windows/Outlook/Exchange/Scheduler acceptance is separate from Linux parser, analyzer and fixture evidence.
