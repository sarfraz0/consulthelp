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

## Production audit, 20 October 2026

**Overall: FAIL / release HOLD.** The [production audit](audit-2026-10-20.md) identifies two new recovery defects: SS-01, missing checkpoint resets an expired partial transfer’s original budget; SS-02, an owned helper survives timeout after its parent exits and cleanup is skipped. See the companion audit for six separate [Embelish findings](../embelish/audit-2026-10-20.md).

**68/68 previous Sendshards scenarios pass** against unchanged production function definitions. Three new failure/destination scenarios yield **1 PASS, 2 FAIL**: a real relative filesystem destination works, while aged-checkpoint loss permits an additional real local push and an exited-parent process fixture leaves its helper alive after GitTimeout. Total: **71 functional scenarios, 69 PASS and 2 FAIL**. Tests restrict Git transport to local files, retain the Windows entry guard and substitute no writer/JSON/native-budget prototype. Owned fixture cleanup targets its recorded child only.

PowerShell 7.6.6/.NET 10.0.12, Git 2.52.0 and PSScriptAnalyzer 1.24.0 on Debian Linux: **0 parser errors; 0 default plus 5.1 syntax/command/type analyzer findings**. AST guideline measurements cover 64 functions plus entry: Read-TransferJson complexity 12/nesting 4, Test-JsonElement complexity 11 and Assert-ControlledDirectory nesting 4 exceed defaults. Entry is complexity 9/nesting 3/66 token-bearing body lines. Count if/elseif, loops, non-default switch clauses, catches, Boolean And/Or; nesting includes if/loop/switch/try containers. Body lines exclude declarations/parameters/comments; nested function scopes are counted separately and public CLI parameter count is exempt. Review/document exceptions rather than claiming a clean complexity gate.

Production code remains byte-identical to commit 07482b5. Current source hashes and detailed reproduction boundaries are recorded in the audit. Windows PowerShell 5.1/Git-for-Windows, native owned process-tree cleanup, durable restart under disk-full/access-denied/hard termination, approved signing/credentials/proxy, Windows filesystem and scheduler/endpoint acceptance remain **NOT RUN**. Dedicated duplicate-code/secret scanner tools were unavailable.

## Production corrections, 20 October 2026 — version 1.0.2

SS-01/SS-02 are implemented. The prior audit's source hashes and failed outcomes describe version 1.0.1 and remain historical. Current synthetic verification uses Debian Linux, PowerShell 7.6.6/.NET 10.0.12, Git 2.52.0 and PSScriptAnalyzer 1.24.0.

| Check | Result |
|---|---|
| Existing protocol suites: 35 + 11 + 12 + 10 cases | 68 PASS |
| Added checkpoint/start/history/native-cleanup-result cases | 10 PASS |
| Parser and default plus explicit Windows PowerShell 5.1 syntax/command/type profile | 0 findings |
| Embedded Windows native adapter compilation with C# 5 and four x64 native structure layouts | PASS; Windows invocation NOT RUN |
| Native help, XML and original UTF-8/no-BOM/LF format | PASS |

Tests retain the real JSON writer/reader, planning, hashing, index/commit/history, pacing, source guards, checkpoint reconciliation and actual local Git pushes. The changed native launch boundary is replaced **only in development AST-loaded tests** by a Linux System.Diagnostics.Process adapter. Relocation checks likewise exercise copies with no docs/XML and reload this development process adapter. No runtime test switch, module, Windows-guard bypass or alternate production launch path was added. These 78 passing cases establish local protocol behavior, not Windows process containment.

New cases verify that the formerly expired partial transfer cannot send the second shard after checkpoint loss; a surviving original-start record rejects missing state before publication; genuine first Push persists matching state/start and completes; legacy state retains its original start on upgrade; conflicting starts are refused; a local unpushed transfer commit cannot manufacture new state; a relative local destination remains consistent; uncertain cleanup prevents an otherwise successful native result; new Plan refuses a surviving start filename; and first-Push index refusal leaves no orphan start record.

Both embedded native sources compile with C# language version 5. x64 layouts match STARTUPINFOEX (112 bytes), JOBOBJECT_EXTENDED_LIMIT_INFORMATION (144), JOBOBJECT_BASIC_ACCOUNTING_INFORMATION (48), SECURITY_ATTRIBUTES (24) and the companion FILE_ID_INFO (24). This does not exercise Windows marshalling, CreateProcess/job APIs, job nesting, endpoint allowlisting or helper behavior. The old Linux exited-parent fixture cannot validate the replacement Windows implementation and is **NOT RUN against that adapter**; it is not silently counted as passed.

**Windows release gate remains open:** real PowerShell 5.1 CLI Plan/Push/resume with no XML, supported Git for Windows and approved hooks/signing/credentials/proxy; parent-exit/inherited-pipe/grandchild timeout and Ctrl+C/hard-host termination; no surviving owned or terminated unrelated processes; job-policy refusal; disk-full/permission/state-loss interruption; UNC/ACL/cloud/linked-worktree and scheduler acceptance. Completed-resume and immutable-budget contracts remain distinct. Runtime-data custody is required if all start/checkpoint evidence is lost. See [recovery](operations.md) and [native-boundary exceptions](implementation-notes.md).

Current implementation SHA-256: `402cf4925c23cad96bdbbf78c7c177c650112c1dec4e362b2e4ea909b7ceda4c`.

## Controlled shell review, 20 October 2026 — version 1.0.3

Scope: standard-user controlled-shell operation, local planning/data and an approved Push workflow. Git ls-remote remains permitted only within Push verification, alongside git push. No networking beyond that workflow, elevation, installation, scheduler or transport change is introduced. Tests use filesystem-only bare remotes with Git restricted to the file transport; no real remote is contacted.

**Reproduced and fixed:** after creation of the exact transfer commit, changing the configured remote caused Send-ItemCommit to publish to the changed destination, then report PushUnverified while inspecting the old URL. The local reproduction confirmed that the second bare repository received the commit. Send-ItemCommit now revalidates the frozen repository, branch and configured destination before remote inspection, attempt checkpointing or push.

Current focused results: **3 PASS** — changed destination leaves both local remotes at the base and does not record an attempt; changed checked-out branch refuses publication; normal local Plan has no transport commands and Push completes with verified frozen bytes. All **78 existing protocol/recovery cases PASS** using the real writer/reader and local Git transactions. Parser/default/5.1 compatibility findings: **0**. Original file encoding/newlines and help XML are preserved. Tests use the previously documented development process adapter on Linux; production Windows process ownership is unchanged by this revision.

The remaining acceptance check for this scope is a real ordinary-user Windows PowerShell 5.1 Plan/Push/resume using approved installed Git, credentials and the Windows Job Object path. That is **NOT RUN** in this Linux environment. No admin access or real production data was used in the review.

Reviewed implementation SHA-256: `a50c90fe321c6cd8d57462577ca5417a0b0a61bb40ba1d88fa66e4e32f868075`.
