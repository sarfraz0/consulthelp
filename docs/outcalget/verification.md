# outcalget verification

[Documentation index](README.md)

## Historical evidence from the original script

The following text is preserved from the embedded help of the original 0.3.5 script. These statements describe earlier revisions and environments; they are not results from this documentation extraction. The original section label was "ENGINEERING / VERIFICATION, 4 OCTOBER 2026".

This revision retains native export semantics and schema-1 parsing; schema 2 is exact
and conditional on sourceType. Pure business preflight and WhatIf remain distinct from
export. ShouldProcess gates parent mutations; operational diagnostic I/O is the disclosed
preview exception. Internal helpers do not introduce independent confirmation flows.

Local complexity/length exceptions remain on the existing parser/validation/native worker/
entry transaction and coherent ownership/recipient/navigation/setup/progress helpers.
Four-space indentation and at most six internal parameters are retained; no claim that
all default complexity/nesting/80-line budgets pass is made. The COM worker combines
acquisition/export/reply/reference lifetime to preserve the common native operation.
No blanket analyzer suppression is used; the narrow Framework Windows.Forms compatibility
exception documents its explicit assembly load. Development tests/fixtures are not shipped.

PASS: PowerShell 7.4.6 parser and PSScriptAnalyzer 1.24.0 default rules plus
PSUseCompatibleSyntax (5.1), PSUseCompatibleCommands and PSUseCompatibleTypes against
win-48_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework, with zero issues
outside the documented narrow explicit Windows.Forms load exception.
PASS (0.3.2 baseline): 285 executable development checks on Linux/PowerShell
7.4.6: 48 portable
regressions; 9 log/scheduler checks; 16 native-settings/protocol/queue-budget checks;
32 event/readiness/ownership checks; 32 baseline CLI checks; 21 lifecycle CLI checks;
58 profile/shared/navigation/path-identity/lease/help checks; 57 profile/concurrent-worker
process checks; 12 setup readiness/feedback/native argument checks. They cover native settings, Gregorian dates, byte-preserved old configs,
named setup/reuse/failure, shared resolution/identity/permissions, safe diagnostics,
output collision/claims/aliases, queued native calls, same-calendar overlap, joined
startup/sync/closure, closed-session setup, setup cancellation/known failure, joined
setup/export lifetimes, shared null-Store synchronization/restoration/failure diagnostics,
and recovery after abrupt worker loss. Child processes, private
transactions and file-sharing handles execute; Windows/Outlook/ACL/connectivity/native
path reads/events are simulated. Argument dispatch executes to a harmless fixture
child; its start-time identity is simulated because it is unavailable in this sandbox.
Actual C# path bridge compilation and native struct
layout are checked; this does not execute kernel32 or prove Windows metadata behavior.
PASS (0.3.3 historical focused): 124 checks rerun on that revision: 64 profile/shared/source-identity
checks and 60 parent/worker process checks. Native source-folder matching, equivalent
entry IDs, shared-source-dependent synthetic ICS contents, and refusal to publish/save
wrong-source results are covered. The complete 0.3.2 suite was not rerun then.
Outlook and Windows boundaries
remain simulated. Parser/analyzer/BOM/CRLF/version/help were checked on the final file.
PASS (0.3.4 historical focused): 130 checks on that revision: 47 synchronization/readiness/ownership
logic checks; 20 native-settings/worker-protocol/queue-budget audit checks; 63 profile/
concurrent parent-worker process checks. Tests exercise interrupted sync recovery,
discarded completion during disconnection, headers-only and Work Offline states,
Internet loss, failed cleanup, three-attempt cap, one original deadline, bounded safe
state observations and malformed-diagnostic refusal. Failed shared export/setup preserves
existing bytes or remains unsaved; recovery diagnostics survive owned Outlook cleanup.
The callback state machine executes with simulated COM events/message pumping; process
fixtures execute real parent/workers and file transactions with Windows/Outlook boundaries
simulated. Their wait fixture supplies chosen sync outcomes, so those process checks
do not establish real Exchange recovery. The complete historical suite was not rerun.
PASS (manual): profile/default compatibility, common native exporter settings, bounded
process/COM ownership and output claims, ShouldProcess, allowlisted diagnostics, declared
side effects and local complexity exceptions. UTF-8 BOM/CRLF checked before saving.
Function/entry complexity was reviewed using AST-guided branch/loop/catch/logical counts
and manual ownership review; no claim of threshold compliance, coverage percentage,
duplication percentage or secret-scanner certification is made.
NOT RUN: Windows PowerShell 5.1 execution, actual Outlook shared/navigation providers,
Exchange freshness/SyncEnd/Quit, locked desktop, Task Scheduler, Windows NLM/ACL/NTFS,
volume-GUID/drive-alias/hard-link behavior on Windows, architecture combinations,
endpoint acceptance or consumer fidelity. No Windows/Outlook environment is available.
USER-REPORTED PASS (0.3.0, one corporate laptop): legacy default export with Classic
initially closed completed successfully. USER-REPORTED FAIL (0.3.0): named Configure
with Classic closed stopped with OUTLOOK_NOT_RUNNING; the attached sanitized log
confirmed that setup incorrectly used the already-open read-only path. Revision 0.3.1
adds owned startup and bounded setup readiness, retains selected-calendar sync before
test export, preserves initial-state diagnostics across setup/export entry, and clarifies
synchronization feedback. Target execution of this fix remains NOT RUN here.
USER-REPORTED FAIL (0.3.1): shared-default selection and readiness succeeded, then
sync failed reading ExchangeStoreType from a null Folder.Store. Microsoft documents
this null value for GetSharedDefaultFolder. Revision 0.3.2 accepts it only for a
validated shared-default Exchange binding; retains native sync/checkpoint/restoration;
records the failure stage before cleanup; and reports sourceType before path validation.
Process fixtures now execute the actual synchronization adapter with null-Store shared
folders instead of bypassing it. Target Windows execution remains NOT RUN here.
Revision 0.3.3 additionally audits the remote lookup/export route, replaces direct
shared owner-ID string comparison with documented MAPI CompareEntryIDs, and checks
CalendarSharing.Folder against the resolved source before any native output write.
Mismatch/unidentifiable sources cannot publish output or complete setup. Tests include
equivalent-ID representations, wrong-source refusals and source-dependent synthetic ICS
payloads. These checks do not prove real provider event completeness or freshness.
USER-REPORTED FAIL (0.3.3): the supplied run reached shared-default lookup, source
readiness and native sync start, then our single-reading connection guard failed in
sync-wait with OUTLOOK_CONNECTION_LOST. No native ICS export was attempted. The record
contains earlier primary mode 700 but omits the actual mode/Offline at failure;
transient recovery versus persistent connection loss cannot be inferred from that log.
Revision 0.3.4 adds bounded recovery with a fresh required sync attempt, records actual
connection transitions on success/failure, and preserves Internet/full-details gates.
This fixes premature failure on recoverable states; it does not claim the unrecorded
state was transient or that the target shared export has now passed. Target Windows,
Exchange and Task Scheduler execution remains NOT RUN here.
Revision 0.3.5 retains an interrupted attempt's event listener until correlated
completion after Stop, before subscribing a replacement. Cleanup uses only the
remaining original synchronization budget and fails closed if completion is absent.
This prevents queued cancellation from contaminating the replacement; cancellation
or other errors on an active attempt still fail. Ten address-free deterministic
regressions pass on Windows PowerShell 5.1, with Outlook events simulated. Actual
shared-calendar export/content acceptance remains unverified for this revision.
No production certification is claimed. Test fixtures, frameworks and additional implementation files are not delivered.

## Documentation extraction, 12 October 2026

PowerShell comments were moved into these documents, with native help supplied by external XML in `docs/outcalget`. The implementation version stays 0.3.5 because its executable logic is unchanged. Source line numbers and the runtime-computed script SHA-256 change as documentation is removed. Historical line numbers refer to the original source and must not be used to locate failures in the extracted script.

Validation used PowerShell 7.6.6 and PSScriptAnalyzer 1.24.0 on Debian Linux. Temporary fixtures and original captures remain outside the repository; no production testing framework or additional runtime dependency was added.

- **PASS:** 38 executable assertions covering parser errors, exact noncomment/nonwhitespace token kind/text equality for both scoped scripts, byte-identical embedded native code, original BOM/newline conventions, native external-help resolution from `/tmp`, all documented parameters, all five example commands/explanations, the synopsis, description, operational notes and vendor links.
- **PASS:** 50 documentation/configuration/CLI assertions covering local links, Claude native import targets, Codex instruction filenames, Copilot filenames and `applyTo` front matter, XML structure and help content, fresh-process `-Help`, `-?` and `-Version` with `-NonInteractive` from `/tmp`, out-of-scope script byte equality, and Git whitespace checks recognizing preserved CRLF. Help examples now distinguish executable commands from explanatory prose.
- **PASS:** Exporter analysis has zero findings using default rules and enabled Windows PowerShell 5.1 compatible syntax/commands/types rules, against the installed `win-48_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework` profile.
- **PASS:** Analyzer regression comparison confirms identical finding rule/severity/message sets before and after extraction for both scripts.
- **FAIL (existing baseline findings):** The Git helper still has two `PSAvoidUsingWriteHost` warnings, unchanged from the original script. No clean analyzer result is claimed for that helper; its presentation behavior was preserved rather than changed during documentation extraction.
- **NOT RUN:** Windows PowerShell 5.1 execution; actual Outlook, Exchange, synchronization, calendar content, Task Scheduler, Windows NLM/ACL/NTFS and endpoint/consumer acceptance. Linux analysis does not establish these outcomes.
- **NOT RUN:** Live GitHub Copilot, Claude Code or Codex instruction-loading sessions. Native configuration filenames, imports and scope metadata were checked structurally.

No real export, Git update helper, calendar publication, task registration or recovery cleanup was executed. Deploy matching external help with the script to preserve detailed help; the ordinary export implementation still runs without Markdown or help data.

## Standalone execution confirmation, 12 October 2026

**PASS:** 43 isolated-copy/missing-XML assertions across all repository scripts, using PowerShell 7.6.6 on Debian Linux. Each script copy was byte-identical and placed alone in a directory with its referenced documentation directory absent. Native `-?` succeeded for the exporter, Embelish and Sendshards with generated syntax and no missing-help errors; the exporter's `-Version` and `-Help` also returned successfully. Embelish and Deembelish business previews completed with an installed 7-Zip path, synthetic inputs and no XML; normal diagnostic output worked, with no archive/extraction data written. Sendshards reached its expected Windows-prerequisite failure instead of a documentation failure.

The calendar Git wrapper was reviewed statically rather than executed: it has no documentation dependency and retains its existing requirement for adjacent `outcalget.ps1`. No real archive/password operation, Git publication or Outlook export was executed. Windows/Outlook/Git-for-Windows business acceptance remains **NOT RUN**. XML is optional help data; executable logic was not changed to establish this requirement.

## Location-independent wrapper, 19 October 2026

`update-calendar.ps1` now uses the exporter’s published absolute output path, resolves the corresponding worktree, and performs literal single-calendar Git operations with `git -C`. It no longer changes location or deletes the prior calendar. Optional `-ExporterPath`, `-Config`, `-Profile`, and native ShouldProcess support cover relocated deployment and explicit configuration without adding runtime documentation dependencies. The exporter implementation and version 0.3.5 remain unchanged.

**PASS:** five synthetic calendar fixture families covering root/`scripts/`/external deployment from unrelated working directories, exact Unicode/literal calendar publication, unrelated staged-file preservation, unchanged output, failure/preview behavior, config/profile and nonadjacent companion resolution, native hook/push rejection, and output outside a repository. See [wrapper documentation](update-calendar.md) for boundaries. Actual Git commands target only filesystem bare remotes; the Outlook/Windows host boundary is a fixture. Native help succeeds with detailed optional XML or generated syntax without it.

**PASS:** twelve separate missing-XML location scenarios across the exporter, Sendshards and both Embelish preview modes. Exporter `-Version`, `-Help`, `-?` and unchanged caller location were checked at a repository root, under `scripts/`, and in an external tools directory. Static review confirms exporter state/output paths are independent of script location and worker launch uses its absolute `$PSCommandPath`. Full Outlook export/worker execution remains **NOT RUN**.

All four scripts have zero parser errors and zero default plus explicit Windows PowerShell 5.1 syntax/command/type analyzer findings using PowerShell 7.6.6/PSScriptAnalyzer 1.24.0 on Debian Linux. Windows/Outlook/Git-for-Windows, signing/credentials, scheduler and endpoint acceptance remain **NOT RUN**. Temporary fixtures live outside the repository; no real user data or external publication was used.
