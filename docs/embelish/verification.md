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

## Production audit, 20 October 2026

**Overall: FAIL / release HOLD.** The [production audit](audit-2026-10-20.md) records six grouped findings (EB-01–EB-06), distinguishing already documented unsafe behavior from new cases. Focused function checks: **3 PASS, 7 FAIL**. A real interactive 7-Zip 25.01 encrypted four-volume archive/base64/restore round trip passed, with SHA-256-identical restored files, nested Unicode names and an empty directory. Three real negative archive/restore checks failed their safety expectations: Linux leading-dot zero-output success, late collision overwrite without Force, and corrupted restore destroying previous output. A real CLI missing-source privacy check also failed. Total: **15 logical functional scenarios, 4 PASS and 11 FAIL**.

The initial suspected native-start stale-exit-code path was not reproduced: the actual adapter and complete operation correctly rejected an unlaunchable executable on this runtime. Those are passing refusal checks, separate from the confirmed leading-dot zero-output defect. The observed native creation password prompt occurred once on this installed Linux build, so the older two-prompt text is not established cross-platform behavior.

PowerShell 7.6.6/.NET 10.0.12 on Debian Linux: **0 parser errors; 0 default plus 5.1 syntax/command/type analyzer findings**. AST measurements cover 15 functions plus the entry; all are within the declared counting defaults, with Invoke-EmbelishOperation at complexity 10/nesting 3/80 token-bearing body lines and entry at complexity 4/nesting 2/43 lines. Public CLI parameter count is exempt. Code remains byte-identical to commit 07482b5; no bug fix or runtime dependency is included in this audit.

All native archive/password work used a synthetic fixture and the actual terminal; isolated functions were AST-loaded, not whole-script dot-sourced. Windows PowerShell 5.1/7-Zip, Windows filesystem/console/application-control/Scheduler acceptance, large-resource/fault-injection boundaries and native lifetime under cancellation remain **NOT RUN**. Dedicated secret/duplicate-code scanners were unavailable. Existing passing extraction/static evidence does not supersede this audit’s failed behavior gates.

## Production corrections, 20 October 2026 — version 1.2.0

EB-01–EB-06 are implemented; the baseline audit above remains historical evidence. Current checks use Debian Linux, PowerShell 7.6.6/.NET 10.0.12, native 7-Zip 25.01 and PSScriptAnalyzer 1.24.0, with synthetic data only.

| Check | Result |
|---|---|
| Original focused safety/refusal regressions | 10 PASS (previously 3 PASS / 7 FAIL) |
| Added publication, recovery, alias, base64, preview and permission scenarios | 16 PASS |
| Native interactive standalone encrypted/split/base64/restore scenarios | 8 PASS |
| Parser, default analyzer and explicit Windows PowerShell 5.1 syntax/command/type profile | 0 findings |
| Original encoding/newlines | UTF-8 without BOM and LF preserved |
| Native help XML and 15 existing helper help entries from unrelated location | PASS |
| Embedded Windows adapter compilation as C# 5; FILE_ID_INFO x64 layout | PASS; Windows invocation NOT RUN |

The native scenarios execute a relocated script without docs/XML, create four leading-dot encrypted volumes, select one set among two, compare all restored SHA-256 hashes/Unicode names/empty directory, reject corrupted data and wrong passwords without damaging previous output, preserve an actual late write during the password prompt, retain a successful Force backup/unrelated file, redact private CLI markers, and refuse ambiguous directory selection. The eight logical scenarios include several outcome assertions. Production password handling remains the native terminal flow, without credential arguments/environment/configuration.

Added publication checks inject destination changes and a final rename conflict, verify both old output and late writers survive, require manual inspection of retained transactions, test exclusive tool locking, check symlink/overlap refusal, verify streaming base64 bytes and no BOM/newline, reject empty output/native warnings, and verify 0700 private Unix staging/new output plus existing file/directory mode preservation. Windows access-DACL cloning is implemented but **NOT RUN** on Windows.

Development PTY trials initially sent synthetic passwords before 7-Zip finished configuring console input; two trials reached the harness's 60-second ceiling. Prompt synchronization was corrected and the native suite rerun. Old output remained intact in the interrupted trial. These development timeouts do not establish a native codec timeout defect. Embelish still has no internal native timeout; cancellation and hard termination require recovery inspection.

**Target release gate remains open:** Windows PowerShell 5.1/7-Zip console execution, native file-ID/short-name/UNC/link behavior, DACL snapshots/managed endpoint policy, interruption at each rename boundary, disk-full/access-denied/resource limits and actual intended source/recipient inventory acceptance. Linux function/native checks and static compatibility do not certify these. Publication retains a rename gap, previous trees, permission-snapshot semantics and documented metadata limits; do not treat it as a complete filesystem backup or authenticity protocol. See [operations](operations.md) and function-named [engineering exceptions](implementation-notes.md).

Current implementation SHA-256: `4917af4c6a2e517aeb9db6115aefabed038597a19be4d4bb91be31388114ae03`.

## Controlled shell review, 20 October 2026 — version 1.2.1

Scope: interactive standard-user operation with installed dependencies and local data. No runtime installation, elevation, scheduler work, new protocol or network call was added. All verification used local synthetic files/repositories. PowerShell 7.6.6/.NET 10.0.12, 7-Zip 25.01 and PSScriptAnalyzer 1.24.0 were already available.

Four focused corrections:

- **Reproduced and fixed:** FromBase64Transform accepted incomplete groups (`Q`, `QQ`, `QQ=`, `QQ==A`, `QUJDRA`), silently dropping incomplete bytes. Bounded strict decoding now rejects them and data after padding, including across read buffers.
- **Reproduced and fixed:** an initialization/protection failure left a newly created empty transaction blocking retry. That invocation's staging is now removed and its lock released; preexisting recovery data is preserved.
- **Confirmation defect fixed:** mandatory Set-Acl inherited confirmation and could be independently skipped after business approval. Its call now explicitly disables that extra confirmation, while permission failures remain fatal. A ShouldProcess stand-in verifies the actual command with an inherited Low confirmation preference in a noninteractive test host; native Windows ACL application is NOT RUN.
- **Static Windows path defect fixed:** stripping the separator before root validation could reinterpret `C:\` as `C:`. Root refusal now precedes trimming/native identity. A local root-refusal check confirms this ordering; Windows drive/UNC execution is NOT RUN.

Current checks: **8 focused cases PASS**, **10 existing focused cases PASS**, **16 publication/permission cases PASS**, **8 native encrypted standalone cases PASS**. Valid/whitespace-wrapped bytes match native .NET conversion across buffer boundaries; malformed restore fails before native password work, releases its output handle and preserves the previous tree. Native encrypted restore checks compare file hashes, Unicode paths, empty directories, corrupt/wrong-password preservation, late collisions and retained Force backups. Parser/default/5.1 compatibility findings: **0**. The original 15 helper help entries load. Native help XML and original UTF-8/no-BOM/LF conventions are preserved.

A real ordinary-user Windows PowerShell 5.1 session with the installed 7-Zip/ACL/file-identity paths is still **NOT RUN** here. Static compatibility and Linux native tests do not establish that acceptance. No broader unattended or filesystem backup certification is claimed.

Reviewed implementation SHA-256: `f9dcd2f8f0af6786ad6e6a7b47c1ea9b150fa2a742f6418d1143673230f866a9`.
