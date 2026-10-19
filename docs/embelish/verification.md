# embelish verification

[Documentation index](README.md)

## Historical evidence from the original script

The following evidence is preserved from embelish 1.1.0 before documentation extraction. It describes earlier development environments and checks, not validation executed in this task. The original section label was "Verification (revision 1.1.0):".

Environment used: PowerShell 7.4.6 (Core) on macOS arm64; 7-Zip (7zz)
    26.03. Windows PowerShell 5.1 and Windows 7z.exe were NOT available in
    the development sandbox.
  PASS         - Script parses (PowerShell AST parser, 0 errors).
  PASS         - PSScriptAnalyzer: 0 errors, 0 warnings (settings: default).
  PASS         - Embelish/Deembelish round-trip through this script over a
                 pseudo-terminal (password entered at the 7-Zip prompt);
                 restored tree is byte-identical to the source, including
                 nested directories and Unicode names.
  PASS         - 25 MB source produced three 10 MB base64 '.txt' shards
                 (single-line ASCII); shards base64-decoded back to volumes
                 and reassembled. Deembelish also works when -SourcePath
                 points at a single shard file. No staging left behind.
  PASS         - Maximum compression (-mx=9, the default) applied: 20 MB of
                 repeated text collapsed to a single 8 KB shard. Each staged
                 binary volume is deleted as its shard is written, so the
                 destination holds only '.txt' shards and no temp files
                 remain.
  PASS         - -WhatIf performs no mutation and reports a preview record.
  PASS         - Exit codes: missing source (2), existing shards without
                 -Force (2), no shards found (2), wrong password (3),
                 -Force re-embelish (0).
  NOT RUN      - Windows PowerShell 5.1 runtime behaviour, Windows console
                 password prompt and Task Scheduler execution (no Windows
                 host in the sandbox).
  NOT APPLICABLE - Code signing (performed by the owner at deployment).

## Documentation extraction, 12 October 2026

Executable tokens, public parameters and implementation version 1.1.0 are preserved. Documentation comments were moved to `docs/embelish` and native help supplied externally. Source line numbers and file hashes change after comment removal. Original line numbers in implementation notes refer to the pre-extraction source.

Validation used PowerShell 7.6.6, PSScriptAnalyzer 1.24.0 and Python on Debian Linux. Temporary captures, fixtures and check scripts remain outside the repository. No archive operation, password prompt, transfer plan against real data, Git push or scheduler registration was executed.

- **PASS:** 97 assertions across both remaining scripts covering parser errors, exact executable token kind/text equality, unchanged string literals and original BOM/newline conventions, functional-only comment directives, native script help, all documented parameters/examples/links, original operational and historical notes, and all 15 Embelish helper help blocks. Explicit documentation corrections concern archive-encryption wording and native password/scheduler behavior; executable behavior is unchanged.
- **PASS:** 99 documentation/configuration/CLI assertions covering repository-wide local links, Codex instruction files, native Claude imports, Copilot scope front matter, MAML structure and public help coverage, fresh-process noninteractive `-?` from `/tmp`, complete four-script inventory, previously staged unrelated-file hashes, and CRLF-aware Git whitespace checks. Native help does not enter archive or transfer work.
- **PASS:** Default and Windows PowerShell 5.1 syntax/command/type analyzer findings are unchanged from each script's captured baseline, using the installed `win-48_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework` compatibility profile.
- **PASS:** Embelish has zero current analyzer findings. Its original 15 helper help blocks resolve through the shared external XML after loading only function definitions with the original source filename.
- **NOT RUN:** Windows PowerShell 5.1 runtime, Windows filesystem/process behavior, interactive Windows/7-Zip password handling, Task Scheduler, real archive round trips, approved remote transfer/resume, endpoint/proxy/DLP and consumer acceptance. Linux static/help checks do not establish these outcomes.
- **NOT RUN:** Live GitHub Copilot, Claude Code and Codex instruction-loading sessions; recognized filenames, imports and scope metadata were validated structurally.

Current versions remain unchanged because executable logic is unchanged. Deploy matching XML help with the scripts to retain detailed help; execution requires no Markdown or help-data file. Historical verification above describes earlier environments, not this extraction.

## Standalone execution confirmation, 12 October 2026

**PASS:** 43 isolated-copy/missing-XML assertions across all repository scripts, using PowerShell 7.6.6 on Debian Linux. Each script copy was byte-identical and placed alone in a directory with its referenced documentation directory absent. Native `-?` succeeded for the exporter, Embelish and Sendshards with generated syntax and no missing-help errors; the exporter's `-Version` and `-Help` also returned successfully. Embelish and Deembelish business previews completed with an installed 7-Zip path, synthetic inputs and no XML; normal diagnostic output worked, with no archive/extraction data written. Sendshards reached its expected Windows-prerequisite failure instead of a documentation failure.

The calendar Git wrapper was reviewed statically rather than executed: it has no documentation dependency and retains its existing requirement for adjacent `outcalget.ps1`. No real archive/password operation, Git publication or Outlook export was executed. Windows/Outlook/Git-for-Windows business acceptance remains **NOT RUN**. XML is optional help data; executable logic was not changed to establish this requirement.

## Relocation correction, 19 October 2026

Revision **1.1.1** fixes caller-relative `-LogDirectory` resolution. Using PowerShell 7.6.6 on Debian Linux, **six Embelish/Deembelish previews passed** across copies at a repository root, `scripts/`, and an external tools directory, with no XML/help/documentation present. PowerShell location was explicitly changed away from the process directory; relative literal source, destination and log paths resolved there. Installed 7-Zip was supplied, diagnostic JSON Lines were checked, and no archive/extraction output was created. Six additional exporter/Sendshards standalone help/version/platform checks passed across the same layouts. A separate invalid-log-path check confirmed that path-resolution failure remains a nonfatal diagnostic warning.

All four repository scripts parse with zero errors and have zero default plus Windows PowerShell 5.1 syntax/command/type analyzer findings. No Windows/7-Zip password session, real archive round trip or scheduler/endpoint acceptance was executed in this task. The earlier archive evidence above remains historical.
