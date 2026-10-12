# sendshards verification

[Documentation index](README.md)

## Historical evidence from the original script

The following evidence is preserved from sendshards 1.0.0 before documentation extraction. It describes earlier development environments and checks, not validation executed in this task. The original section label was "Verification for revision 1.0.0:".

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

## Documentation extraction, 12 October 2026

Executable tokens, public parameters and implementation version 1.0.0 are preserved. Documentation comments were moved to `docs/sendshards` and native help supplied externally. Source line numbers and file hashes change after comment removal. Original line numbers in implementation notes refer to the pre-extraction source.

Validation used PowerShell 7.6.6, PSScriptAnalyzer 1.24.0 and Python on Debian Linux. Temporary captures, fixtures and check scripts remain outside the repository. No archive operation, password prompt, transfer plan against real data, Git push or scheduler registration was executed.

- **PASS:** 97 assertions across both remaining scripts covering parser errors, exact executable token kind/text equality, unchanged string literals and original BOM/newline conventions, functional-only comment directives, native script help, all documented parameters/examples/links, original operational and historical notes, and all 15 Embelish helper help blocks. Explicit documentation corrections concern archive-encryption wording and native password/scheduler behavior; executable behavior is unchanged.
- **PASS:** 99 documentation/configuration/CLI assertions covering repository-wide local links, Codex instruction files, native Claude imports, Copilot scope front matter, MAML structure and public help coverage, fresh-process noninteractive `-?` from `/tmp`, complete four-script inventory, previously staged unrelated-file hashes, and CRLF-aware Git whitespace checks. Native help does not enter archive or transfer work.
- **PASS:** Default and Windows PowerShell 5.1 syntax/command/type analyzer findings are unchanged from each script's captured baseline, using the installed `win-48_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework` compatibility profile.
- **FAIL (existing baseline findings):** Sendshards retains 15 analyzer warnings: 7 `PSUseShouldProcessForStateChangingFunctions`, 4 `PSUseSingularNouns`, 2 `PSUseDeclaredVarsMoreThanAssignments`, 1 `PSUseOutputTypeCorrectly`, and 1 `PSAvoidAssignmentToAutomaticVariable`. No new findings or suppressions were introduced; a clean analyzer result is not claimed. Fixing these is separate implementation work.
- **NOT RUN:** Windows PowerShell 5.1 runtime, Windows filesystem/process behavior, interactive Windows/7-Zip password handling, Task Scheduler, real archive round trips, approved remote transfer/resume, endpoint/proxy/DLP and consumer acceptance. Linux static/help checks do not establish these outcomes.
- **NOT RUN:** Live GitHub Copilot, Claude Code and Codex instruction-loading sessions; recognized filenames, imports and scope metadata were validated structurally.

Current versions remain unchanged because executable logic is unchanged. Deploy matching XML help with the scripts to retain detailed help; execution requires no Markdown or help-data file. Historical verification above describes earlier environments, not this extraction.


The original adjacent `#requires` and help block did not expose detailed help to native lookup. Separating the external-help directive from `#requires` restores discovery. The original help text was compared using a development-only copy with that separator, without executing any business logic; the captured original executable tokens remain the unchanged implementation baseline.

## Standalone execution confirmation, 12 October 2026

**PASS:** 43 isolated-copy/missing-XML assertions across all repository scripts, using PowerShell 7.6.6 on Debian Linux. Each script copy was byte-identical and placed alone in a directory with its referenced documentation directory absent. Native `-?` succeeded for the exporter, Embelish and Sendshards with generated syntax and no missing-help errors; the exporter's `-Version` and `-Help` also returned successfully. Embelish and Deembelish business previews completed with an installed 7-Zip path, synthetic inputs and no XML; normal diagnostic output worked, with no archive/extraction data written. Sendshards reached its expected Windows-prerequisite failure instead of a documentation failure.

The calendar Git wrapper was reviewed statically rather than executed: it has no documentation dependency and retains its existing requirement for adjacent `outcalget.ps1`. No real archive/password operation, Git publication or Outlook export was executed. Windows/Outlook/Git-for-Windows business acceptance remains **NOT RUN**. XML is optional help data; executable logic was not changed to establish this requirement.
