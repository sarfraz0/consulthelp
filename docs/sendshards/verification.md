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

## Verification and gap audit, 18 October 2026

**Overall: FAIL.** The [detailed audit](audit-2026-10-18.md) records six reproduced findings: initial planning-budget overload failure, checkpoint-replacement failure, escaped duplicate-key acceptance, missing manifest string-type checks, the no-work index-policy gap, and PowerShell 6/7 timestamp round-trip incompatibility. The PowerShell-version issue is explicitly separated from the unexecuted Windows PowerShell 5.1 baseline.

Executed on Debian Linux with PowerShell 7.6.6, .NET 10.0.12, Git 2.52.0 and PSScriptAnalyzer 1.24.0: **58 functional scenarios, 44 PASS and 14 FAIL**, plus two passing standalone CLI smoke checks. Repeated failing scenarios map to six root causes. Downstream fixtures use finite deadlines to isolate the initial planning failure; two passing scenarios and the no-work-index reproduction use a scratch-only writer correction. Those are not clean end-to-end passes of the unchanged production script.

Parser errors: **0**. In-file functions inventoried: **63**. Unresolved statically named PowerShell commands in the available environment: **0**. Default and explicit 5.1 syntax/command/type analysis retains the same **15 warnings**; the analyzer gate remains FAIL. Exact executable-token comparison confirms all findings concern code already present before documentation extraction.

Only synthetic repositories and filesystem-only bare destinations were used, with Git restricted to the file transport. No real transfer destination or corporate data was used. Windows runtime, owned process-tree cleanup, interruption/durability, actual multi-file pacing, signing/credentials/proxy/Scheduler and endpoint acceptance remain NOT RUN or incomplete as detailed in the audit matrix. Production code is unchanged; this entry supersedes any interpretation of the earlier extraction checks as business-operation verification.

## Fixes and relocation verification, 19 October 2026

Revision **1.0.1** fixes F1–F6 from the [18 October audit](audit-2026-10-18.md). Executed on Debian Linux with PowerShell 7.6.6, .NET 10.0.12, Git 2.52.0 and PSScriptAnalyzer 1.24.0:

| Executed checks | Result |
|---|---|
| Original 35 function/native-Git regressions | 35 PASS |
| Original 11 JSON/manifest/checkpoint/transfer scenarios, with prototype override removed | 11 PASS |
| Original 12 budget/size/stream/manifest/history guards | 12 PASS |
| Added adversarial validation, actual paced transfer, persisted resume and relocation scenarios | 10 PASS |
| All four scripts: parser and default plus 5.1 syntax/command/type compatibility analysis | 0 errors; 0 analyzer findings |

All 68 Sendshards functional scenarios use the production function definitions loaded with the AST, without executing or altering the Windows entry-point guard. Fixture planning contexts now start with the actual null deadline and integer timeout; there is no writer prototype or function-body rewrite. One guard explicitly assigns a finite deadline to test monotonic expiration, and a hash-array assertion now expects the earlier strict-type rejection (`JsonString`). Temporary fixtures/captures stay outside the repository.

Added checks include nested escaped/case-colliding keys, native item-name mapping, rejection of hidden type metadata, exact timestamp-string preservation and refusal of noncanonical dates, malformed null/array/object/Boolean/numeric manifest and checkpoint identities, a real two-file transfer with at least five seconds pacing, saved completed/no-work resumes with foreign-index refusal, and saved-state resume with the exact next shard already staged. Actual local Git commits/pushes are verified against filesystem-only bare remotes; `GIT_ALLOW_PROTOCOL=file` prevents other Git transports.

Relocated copies, with no referenced XML or other documentation, completed local plan publication → saved-plan reload → actual local commit/push → saved-checkpoint resume from an unrelated PowerShell location. Layouts covered the repository root, `scripts/`, and an external tools directory. Absolute data paths and caller-relative plan resolution were checked. Source data remains in its worktree; plan/state remain outside it. Separate standalone CLI checks reached native help or the expected Windows-platform refusal.

The 15 old analyzer findings are resolved through code corrections and seven narrowly justified per-function ShouldProcess exceptions described in [implementation notes](implementation-notes.md). No general rule disablement was introduced. This is static compatibility evidence, not Windows runtime acceptance.

**NOT RUN:** Windows PowerShell 5.1/Git-for-Windows execution, nested Windows/drive/UNC/junction/linked-worktree behavior, Windows owned process-tree cleanup, abrupt termination/disk-full/access-denied durability, approved signing/credentials/proxy/LFS integration, Task Scheduler and endpoint acceptance. The remaining platform/operational matrix in the original audit still applies. No real shards, corporate data or external transfer destination were used.
