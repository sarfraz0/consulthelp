#requires -Version 5.1

<#
.SYNOPSIS
Exports configured Outlook Classic calendars, with named configurations and coordinated concurrent runs.
.DESCRIPTION
outcalget 0.3.5. One Windows PowerShell 5.1 file; no administrator rights, runtime
modules/downloads, Graph permissions, endpoint exclusions or policy bypass.
The native Outlook exporter preserves Full details, today through today + 29 LOCAL
calendar dates inclusive, unrestricted working hours and saved privacy/attachment
choices. Private details and attachments default to off during new setup.
Existing schema-1 configurations from 0.1.x and 0.2.0 are accepted unchanged.

No -Profile means the existing default config.json. -Profile default selects that
same file. -Profile team selects profiles\team.json beneath the existing local state
root. A missing named profile fails with CONFIG_MISSING, exit 2: no default fallback
and no setup prompt. -Configure -Profile team creates it after a successful test
export. Named setup refuses an existing profile; use a new name to create another
binding. -Config remains available for existing explicit configuration paths and
cannot be combined with -Profile. Export profile names are NOT Outlook MAPI profile
names; the configuration's profileName still binds the actual Outlook session.

Calendar setup offers (1) the original Outlook folder picker, (2) another mailbox's
shared default Exchange calendar by full owner email, and (3) accessible calendars
already added to Outlook's Calendar navigation list. Option 3 addresses shared/linked
calendars absent from PickFolder. Only option 2 uses GetSharedDefaultFolder; it does
not select arbitrary secondary shared calendars. Both paths use the same native
GetCalendarExporter/SaveAsICal operation. Outlook sharing/export permissions apply.
No email, invitation or sharing request is sent.
For validated option-2 shared-default bindings, Outlook can return Folder.Store=null
by design. This is accepted after Exchange owner identity has been revalidated;
native Application Folders synchronization still runs before an owned-session export.
A null Store on an ordinary folder binding or an explicit non-Exchange store fails
cleanly instead of claiming synchronization. Each export additionally inspects the
native CalendarSharing.Folder and uses NameSpace.CompareEntryIDs to verify that the
exporter targets the resolved calendar before SaveAsICal. Shared mailbox identity
pinning also uses CompareEntryIDs because different entry-ID representations can
refer to the same object. Native folder bindings additionally retain store validation.

Internet availability is required, using Windows Network List Manager. When Outlook
Classic is already open by the user, the helper attaches, performs no explicit sync
and leaves it open. Otherwise it starts the machine-registered Classic installation
in the saved MAPI profile for configured exports; interactive setup uses normal Outlook
startup/default-profile behavior and can display its usual profile/sign-in dialogs.
The helper waits for native readiness/online Exchange/group sync and requests normal
Quit only for the exact process it started. Other exports can join
that process. The creator waits for all joined usage leases before closing it; joined
runs synchronize their own selected calendars and never request Quit themselves.
Your ordinary configured Windows user must be logged on in the same desktop session.
-NonInteractive cannot suppress Office sign-in/security dialogs. Logged-off/server
Office automation and new Outlook are unsupported.

REQUIRED TASK SCHEDULER SETTINGS
Create one task per export profile when you want calendars to run concurrently.
General: configured ordinary user; Run only when user is logged on; Run with highest
privileges OFF. Do not use SYSTEM, Password/S4U or Run whether user is logged on or not.
Choose the desired trigger; use a 1-2 minute delay for an at-logon trigger.

Action = Start a program; substitute your approved absolute paths:
  Program/script: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
  Add arguments:  -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -STA -File "C:\ApprovedTools\outcalget.ps1" -Scheduled -Profile team -Debug
  Start in:       C:\ApprovedTools
Omit -Profile team for the original default calendar. Append -Config "C:\...\config.json"
only for a legacy explicit configuration path, never alongside -Profile.
No shell redirects, execution-policy switches, task registration or elevation are used.

For each task: If already running, Do not start a new instance. Different tasks can
run concurrently with different profiles. Do not enable automatic failure retries.
Optional Run as soon as possible after a missed start is suitable only while signed in
and online. Do not require idle; allow battery execution and avoid battery-transition
stops if needed. Do not wake the laptop by default. Sleep/logoff prevents export.
Allow at least 60 minutes before Scheduler's external stop limit with the defaults,
so several bounded coordination, worker and cleanup waits cannot trigger a hard stop.
Increase that limit if you raise budgets or add many contending calendars; no duration
is guaranteed for arbitrary contention. Network condition is optional: the script
always checks Windows Internet status. An employer may restrict task creation.

CONFIGURATION AND CONCURRENCY
State root: %LOCALAPPDATA%\osp\consulthelp\outcalget.
Default: config.json. Named: profiles\<lowercase-name>.json. Names are case insensitive,
1-48 ASCII letters/digits/hyphens/underscores, starting with a letter/digit; reserved DOS
device names are forbidden. Existing configurations are bound to Windows SID, Outlook
MAPI profile and exact native folder identifiers. Shared-default bindings instead pin
the resolved Exchange directory owner ID and save its canonical SMTP address; every
worker resolves and checks that identity again. Mailbox/directory changes can require
explicit reconfiguration. Configuration and operational state contain sensitive IDs.

Different output paths are mandatory. Persistent private output-claims\<hash>.json
records reserve each physical filename for its canonical configuration path. The
managed default/named catalog is also checked, protecting old unregistered configs.
Explicit custom configs register claims when used; other unused custom config paths
cannot be discovered. Copies/moves do not automatically transfer output ownership.
A claim remains if a profile or output file is deleted, preventing accidental reuse.
Failed new setup removes only its own new claim when the outcome is certain.

To deliberately retire a claim: finish all helper runs; review the protected record's
configurationIdentity and outputIdentity; ensure every referencing task/configuration
has been retired or redirected, then remove only that claim. Claims are not recovery
markers and must not be routinely deleted. No automated reassignment is provided.
Catalogs over 256 profiles and malformed managed configurations fail clearly.

Drive aliases are resolved through read-only Windows volume-GUID metadata. Hard-linked
configuration/output/staging targets and reparse points are rejected. Output cannot
be inside protected runs, locks, profiles, output-claims or default logs directories.
The canonical destination is checked again immediately before publication. Existing
conservative local fixed-NTFS/private-ACL/path-length rules are retained (see SECURITY).

Per-configuration, actual-calendar and output locks reject unsafe overlap with exit 6.
Known native folder bindings and shared-owner bindings are reserved before startup.
All Outlook COM worker operations are queued through one bounded gate: the script
supports concurrently active runs; native Outlook calls are serialized. Short lifecycle
coordination gates protect startup, sync membership and close decisions. Live parent
PID/start-time/session checks plus shared usage handles prevent closing under another
export. If the last usage lease cannot be confirmed released within the budget, Outlook
is left open and recovery is required. One Outlook instance supports one MAPI profile;
concurrent export profiles must use that same currently running Outlook MAPI profile.
Install the same current script for all tasks. Upgrade only with no helper running;
do not run older copies concurrently, since they do not implement these leases.
.PARAMETER Configure
Interactive source/options selection and a test export. Saves only after success and
requires an unoccupied output filename. Never put this switch in a scheduled action.
With -Profile, creates a new named configuration and refuses an existing named file.
Starts Classic if closed, waits for its profile/stores and online Exchange, then offers
selection. The selected calendar is synchronized before the test export when the
session is script-owned. Existing user sessions are left open and receive no explicit
sync. Only this run's owned startup process is eligible for normal Quit during cleanup.
.PARAMETER ExportProfile
Alias: -Profile. Names an export configuration; no selector or default reuses the
original config.json. This never selects or changes the Outlook MAPI profile.
.PARAMETER Config
Legacy explicit absolute local .json path; mutually exclusive with -Profile.
.PARAMETER Check
Read-only business preflight of config/source/apparent path/output ownership. Requires
already-open Outlook. No native export, output write, claim registration, explicit sync
or freshness/consumer test. Operational logs, locks and private worker files can occur.
.PARAMETER Scheduled
Non-prompting saved export plus scheduled metadata. Does not create a task or alter
business behavior. Incompatible with Configure, Check, Help and Version parameter sets.
.PARAMETER CoordinationTimeoutSeconds
Default 300 seconds, range 15-1800. Applies to each catalog/lifecycle/COM-queue or
joined-session cleanup wait. Worker queue time is additional to its native-work budget;
parent monitoring detects native start separately and also enforces a total deadline.
A busy same-configuration/calendar/output lock fails immediately rather than queuing.
.PARAMETER StartupTimeoutSeconds
Export-only readiness budget, default 120 seconds, range 15-600. Includes native COM,
saved-source and primary online Exchange readiness for script-owned sessions; blocking
COM work is isolated in a bounded worker process. Configure uses the embedded defaults.
.PARAMETER SyncTimeoutSeconds
Export-only native group synchronization budget, default 180 seconds, range 30-600.
Applies to a process started here and to joined script-owned sessions. Existing user
sessions receive no explicit sync. Native sync completion is not proof of freshness.
Connection recovery shares this same budget, never a fresh timeout. At most three
native starts are allowed; each interrupted attempt is stopped/detached and a new
correlated completion is required after recovery. Explicit Work Offline fails promptly.
.PARAMETER LogDirectory
Absolute approved private local directory; default state-root\logs. Does not relocate
configuration, ownership claims or operational state. Overrides the embedded log path.
.PARAMETER Json
One result JSON envelope on stdout, including handled failures. Schema 1 retains all
existing fields and adds namedProfile, sourceType and joinedOwnedSession. Debug does
not change it. Do not combine with WhatIf/Confirm, whose host output can corrupt JSON.
.PARAMETER Help
Compatibility alias for full Get-Help; also supports Get-Help and -?.
.PARAMETER Version
Prints 0.3.5 without Outlook access or run records.
.EXAMPLE
powershell.exe -NoLogo -NoProfile -STA -File "C:\ApprovedTools\outcalget.ps1" -Configure -Profile team -Debug
Classic starts automatically if closed. Choose option 3 for a calendar already in the Calendar
navigation list, or option 2 for a shared mailbox's default calendar. Choose a distinct
output path. Later runs do not ask for a calendar, owner, options or output path.
.EXAMPLE
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\ApprovedTools\outcalget.ps1" -Scheduled -Debug
Exports the existing default configuration without reconfiguration.
.EXAMPLE
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\ApprovedTools\outcalget.ps1" -Scheduled -Profile team -Debug
Headless named export. Use the same arguments in that profile's Task Scheduler action.
.EXAMPLE
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\ApprovedTools\outcalget.ps1" -Check -Profile team -Json
Read-only preflight. Outlook must already be open. Does not replace a real export test.
.EXAMPLE
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\ApprovedTools\outcalget.ps1" -Scheduled -Profile team -WhatIf -Debug
Preview: no source access, export, startup, output claim or configuration changes.
Minimal run logging may occur; retention is skipped. Configuration/source is not validated.
.INPUTS
None. Configurations are generated JSON data. Private worker switches are not public API.
.OUTPUTS
Normal CLI: UTC phase feedback, export/check summary, sanitized error on stderr, warnings
and run-record path. Native hosts can render information/warnings on stdout. -Json
suppresses phase feedback and provides one result envelope. Nonzero exits distinguish
failure. Output ICS bytes are native Outlook output, not manually constructed events.
.NOTES
RUN RECORDS
Required unique bounded UTF-8 JSON records initialize before business work, checkpoint
parent phases, and finish with completed=true on handled completion. Normal records
include outcome, safe source/lifecycle flags and phases. Debug adds allowlisted host
versions, script SHA-256, codes, HRESULTs and timings to the SAME record, without prompts.
Synchronization diagnostics retain at most 24 changed connection observations: Offline,
primary Exchange mode, elapsed milliseconds and attempt number. These describe the
primary account, not an independent measurement of the selected shared mailbox/cache.
No calendar content, owner addresses, native IDs, profile names, paths, tokens or raw
exceptions/child output enter logs. The additive failurePhase field preserves the
primary error stage even after synchronization restoration or Outlook closure. The result
envelope intentionally gives outputPath and runLogPath. Share only a reviewed run record; do not share config/profile files,
output-claims, runs, partial ICS files or calendar exports for debugging.

Retention attempts 200 closed records and 14 days, protecting active handles; inaccessible
files remain with warnings. Required checkpoint failure stops later work. A completion
log failure after publication reports exit 5 with published=true; a successful export is
not undone. Optional debug enrichment cannot mask the primary result. A missing/incomplete
record means UNKNOWN, including abrupt termination or Scheduler's external stop. An old
calendar.ics may be stale after failure; correlate runId/time/completed/published/exitCode.
Host binding, signing/policy/parse refusal and hard termination can precede any record.

RECOVERY
Only owned child workers can be stopped for timeout; Outlook is never forced closed.
Native work may continue after worker loss. Per-output and global COM .lock.pending
markers plus the owned lifecycle marker block unsafe retries, including other profiles.
Normal completed COM calls remove only their own marker while still holding the COM
queue gate. Handled uncertain work leaves Outlook open and retains protected artifacts.
An alive lifecycle marker with ready/closing phase can belong to normal concurrent work;
its existence alone is not a failure. Never remove a marker while a helper is running.

To recover: close Outlook normally; ensure no outcalget parent/worker remains; inspect
state-root\locks\*.lock.pending and their runId/time. Review runs\<runId>\sync\sync-state.json
before deletion. If settingRestored=false, restore the selected folder checkbox in
Outlook's Application Folders Send/Receive group to originalInclusion, or ask IT to do
so. Multiple concurrent runs can have distinct recovery state; review every retained
run involved. Then remove only the failed run's output/COM/lifecycle pending markers
and, if needed, its runs\<runId> folder and .outcalget-<runId>.partial.ics stage. Leave
ordinary .lock files, output-claims and final calendars intact. Old 0.2.0 lifecycle
markers also require this review; they cannot be adopted as current live sessions.
Rerun with -Debug. Do not clear recovery state while Outlook could still be writing.

SYNCHRONIZATION AND CLOSURE LIMITS
Successful owned/joined runs report synchronization=group_completed and a readable
synchronizationSummary in the run record; CLI success says Outlook synchronization
completed. The backward-compatible freshness=not_verified field means that server/cache
equivalence was not independently checked, not that the observed synchronization was
skipped or failed. Existing user sessions can have stale cached
sources despite Internet status; synchronization=existing_session_not_requested records
that choice. Owned/joined sessions require Offline=false and full-item Exchange mode
600/700/800, include the selected folder temporarily in SyncObjects.AppFolders, observe
correlated SyncStart/OnError/SyncEnd, and restore its original InAppFolderSyncObject value.
Temporary nonready readings wait within the original sync budget. An interrupted
attempt cannot authorize export after recovery: a fresh attempt must complete while
full-item connection is observed. Attempts are capped at three. Work Offline modes
100/200, native sync errors and Internet loss fail promptly; no Outlook connection
settings are changed. Persistent connected headers-only mode 500 is reported as
OUTLOOK_HEADERS_ONLY; persistent other nonready states retain OUTLOOK_CONNECTION_LOST
with the actual last mode/Offline flag. Repeated recovered interruptions report
OUTLOOK_CONNECTION_UNSTABLE. Full connection without completion times out normally.
A narrow C# ComEventsHelper bridge and STA message pump use documented Outlook event
metadata. Only numeric sync errors are retained. Application Folders can be created by
Outlook and other members can sync. Non-Exchange or unavailable sync sources fail clearly.
No Namespace.SendAndReceive or MailItem.Send is called. A group completion cannot prove
that every shared-calendar server change reached the cache. Test representative content.

Closure requires exact started PID/start time, one Outlook process in this desktop,
no open items/additional explorers/inline drafts, released participant leases and a
normal Quit followed by confirmed process exit. User activity or uncertain native work
leaves Outlook open with nonzero status and recovery artifacts. Avoid using a script-
started session during export; the user-activity check-to-Quit interval cannot be atomic.
Abrupt parent termination cannot guarantee closure or setting restoration. Joined-run
sync failure/cleanup uncertainty can also prevent the creator from safely closing.

SECURITY / COMPATIBILITY
Windows PowerShell 5.1 Desktop STA, FullLanguage and Windows .NET Framework 4.5+ are
required by the COM adapter. Existing policy must permit COM and the transparent
Add-Type bridges; blocked initialization fails without fallback/bypass. No runtime
packages, credentials, Graph or tenant consent are needed. The file is unsigned: use
your employer's existing approved signing/deployment process. Worker processes execute
this SAME unchanged file under the same user with NoLogo/NoProfile/NonInteractive/STA.
Startup reads HKLM App Paths without modification. Configured exports pass only the
quoted saved /profile name to Classic; interactive setup supplies no profile switch
and records the running MAPI profile after readiness and calendar selection. No broad Office process kill, registry/ACL/policy change,
endpoint enumeration/exclusion or self-elevation is performed.

Path metadata uses OPEN_EXISTING read-only kernel32 handles, volume-GUID final paths
and file link counts; no backup privilege is enabled. Native Outlook sync and normal
corporate monitoring can perform network activity. Internet gating uses Windows NLM
IsConnectedToInternet, without helper HTTP/ping probes or overrides; corporate Windows
status may report unavailable despite internal service access. This is a clean failure,
not proof of Exchange reachability. Connectivity is checked before business work, during
owned sync, before native save and publication, but can still disappear between checks.

Retained deployment boundary: fixed local NTFS, no UNC/removable/reparse/short-name paths,
absolute paths <=230 characters, output directory <=165 for staging space, private
user/SYSTEM/Administrators ACLs with creator/owner recognized. Unknown corporate ACL
principals fail. These are this tool's conservative snapshot rules, not general Windows
requirements; supporting managed exceptions needs explicit design and target testing.
Output defaults to 25 MiB max (1-250), timeout 120 seconds (30-600). Size is checked after
native export, not a disk quota. ICS validation is structural, not complete RFC/consumer
fidelity. Local dates include earlier appointments today; Outlook handles private details.

EXIT CODES
0 success/check/preview/help/version; 2 configuration/binding/output ownership;
3 host/session/Internet/Outlook/access; 4 native export; 5 storage/logging/recovery;
6 unsafe overlap; 7 bounded timeout/coordination; 8 handled cancellation; 9 protocol/
unexpected error. Pre-body and abrupt-stop failures can use host/Scheduler codes.

UPGRADE AND ACCEPTANCE
Replace only the script while no helper is running; existing configuration and scheduled
default action remain valid. Do not reconfigure a working default just to upgrade.
For a named calendar: run the Configure/Profile example; Classic starts if closed. Select the
correct source/output; Check/Profile preflight, then Scheduled/Profile/Debug actual export.
Confirm current completed=true, published=true, exitCode=0 plus a refreshed destination
and correct full details, private/attachment choices and 30-date range. Inspect recurring,
all-day, Unicode, shared/private and boundary events in the intended calendar consumer.

Test each task with Scheduler's Run, Last Run Result 0x0 AND its corresponding record/file;
then a short temporary trigger while signed in. Test a locked desktop if required. Test
user-open Outlook (left open), initially closed Outlook (started/synced/closed) and no
Internet (failure preserves old file and never starts Outlook). Test two distinct profiles
simultaneously with Outlook both open and closed: files stay separate; the startup creator
closes only after both finish. Test a missing profile and duplicate output refusal. If a
shared navigation binding cannot reopen after Classic restarts, try the shared-default
owner option for default calendars; secondary provider-specific calendars can remain
unavailable and fail explicitly. Do not replace native folder opening with cache scraping.

FUNCTIONAL ACCEPTANCE FOR A SHARED CALENDAR
A completed synchronization group or exit 0 alone does not prove server freshness
or event completeness. After setup, inspect the actual generated .ics for a known
appointment unique to that shared calendar within the next 30 local dates, checking
its date/time and details. Compare with a manual full-details Outlook export using
the same window; include recurring and all-day appointments when present. A valid
empty file is allowed for a genuinely empty window, so check expected content rather
than treating a nonzero event count as a source guarantee. Development fixtures
exercise source routing and publication, not real Exchange event data.

ENGINEERING / VERIFICATION, 4 OCTOBER 2026
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
.LINK
https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.exchangeconnectionmode
.LINK
https://learn.microsoft.com/en-us/office/vba/api/outlook.olexchangeconnectionmode
.LINK
https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.compareentryids
.LINK
https://learn.microsoft.com/en-us/office/vba/api/outlook.calendarsharing.folder
.LINK
https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.store
.LINK
https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getshareddefaultfolder
.LINK
https://learn.microsoft.com/en-us/office/vba/api/outlook.navigationfolder.folder
.LINK
https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.logon
.LINK
https://learn.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-logontype-principaltype-element
#>
[CmdletBinding(PositionalBinding = $false, SupportsShouldProcess = $true, ConfirmImpact = 'Low', DefaultParameterSetName = 'Export')]
param(
    [Parameter(ParameterSetName = 'Configure')]
    [switch]$Configure,
    [Parameter(ParameterSetName = 'Check')]
    [switch]$Check,
    [Parameter(ParameterSetName = 'Help')]
    [switch]$Help,
    [Parameter(ParameterSetName = 'Version')]
    [switch]$Version,
    [Parameter(ParameterSetName = 'Export')]
    [switch]$Scheduled,
    [Parameter(ParameterSetName = 'Export')]
    [ValidateRange(15, 600)]
    [int]$StartupTimeoutSeconds = 120,
    [Parameter(ParameterSetName = 'Export')]
    [ValidateRange(30, 600)]
    [int]$SyncTimeoutSeconds = 180,
    [Parameter(ParameterSetName = 'Export')]
    [Parameter(ParameterSetName = 'Check')]
    [switch]$Json,
    [Parameter(ParameterSetName = 'Export')]
    [Parameter(ParameterSetName = 'Check')]
    [Parameter(ParameterSetName = 'Configure')]
    [string]$Config,
    [Parameter(ParameterSetName = 'Export')]
    [Parameter(ParameterSetName = 'Check')]
    [Parameter(ParameterSetName = 'Configure')]
    [Alias('Profile')]
    [string]$ExportProfile,
    [Parameter(ParameterSetName = 'Export')]
    [Parameter(ParameterSetName = 'Check')]
    [Parameter(ParameterSetName = 'Configure')]
    [ValidateRange(15, 1800)]
    [int]$CoordinationTimeoutSeconds = 300,
    [Parameter(ParameterSetName = 'Export')]
    [Parameter(ParameterSetName = 'Check')]
    [Parameter(ParameterSetName = 'Configure')]
    [string]$LogDirectory,
    [Parameter(ParameterSetName = 'Worker', DontShow = $true)]
    [switch]$Worker,
    [Parameter(ParameterSetName = 'Worker', DontShow = $true)]
    [string]$WorkDirectory,
    [Parameter(ParameterSetName = 'Worker', DontShow = $true)]
    [string]$WorkerToken
)
# -Debug is PowerShell's common parameter. Consume it without Write-Debug/Inquire prompts.
$script:DiagnosticsRequested = $PSBoundParameters.ContainsKey('Debug') -and [bool]$PSBoundParameters['Debug']
$DebugPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Set-StrictMode -Version 3.0

function Stop-Outcalget {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Raises a classified exception only; no state mutation.')]
    [CmdletBinding()]
    param([string]$Code, [string]$Message, [int]$ExitCode = 9, [Exception]$InnerException = $null)
    $e = [InvalidOperationException]::new($Message, $InnerException)
    $e.Data['OutcalgetCode'] = $Code
    $e.Data['OutcalgetExit'] = $ExitCode
    # Optional phase metadata can be absent in a helper-only host before entry initialization.
    $e.Data['OutcalgetPhase'] = Get-Variable -Name Phase -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    throw $e
}

function Get-ExceptionInfo {
    param($Record)
    $e = $Record.Exception
    $code = 'INTERNAL_ERROR'
    $exitCode = 9
    $message = 'The helper encountered an unexpected failure. Review the run record; add -Debug for detail.'
    $types = [Collections.Generic.List[string]]::new()
    $hresult = $null
    $failurePhase = $script:Phase
    for ($i = 0; $null -ne $e -and $i -lt 8; $i++) {
        $types.Add($e.GetType().FullName)
        if ($e.Data.Contains('OutcalgetPhase')) {
            $failurePhase = [string]$e.Data['OutcalgetPhase']
        }
        if ($e.Data.Contains('OutcalgetCode')) {
            $code = [string]$e.Data['OutcalgetCode']
            $exitCode = [int]$e.Data['OutcalgetExit']
            $message = $e.Message # Only messages constructed by Stop-Outcalget.
        }
        if ($e.HResult -ne 0) {
            $bits = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$e.HResult), 0)
            $hresult = '0x{0:X8}' -f $bits
        }
        $e = $e.InnerException
    }
    if ($code -eq 'INTERNAL_ERROR') {
        if ($failurePhase -in @('ical-save', 'exporter-create', 'exporter-settings', 'exporter-source-validate')) {
            $code = 'OUTLOOK_EXPORT_FAILED'
            $exitCode = 4
            $message = 'Outlook could not export the calendar. Check Outlook for a dialogue and use -Debug for the failure code.'
        }
        elseif ($failurePhase -eq 'outlook-attach') {
            $code = 'OUTLOOK_NOT_READY'
            $exitCode = 3
            $message = 'Cannot attach to Outlook Classic. Open it normally, finish sign-in, switch focus back here and try again.'
        }
        elseif ($failurePhase -eq 'source-validate') {
            $code = 'SOURCE_UNAVAILABLE'
            $exitCode = 2
            $message = 'The configured Outlook source is unavailable. Check the running profile; use -Configure if the binding changed.'
        }
        elseif ($failurePhase -in @('outlook-launch', 'outlook-ready-wait', 'outlook-setup-wait', 'outlook-online-wait', 'sync-prepare', 'sync-start', 'sync-wait', 'sync-connection-wait', 'sync-connection-recovered', 'sync-restore')) {
            $code = 'OUTLOOK_LIFECYCLE_FAILED'
            $exitCode = 3
            $message = 'Outlook startup, readiness or synchronization failed. Inspect Outlook interactively and review the debug run record.'
        }
        elseif ($failurePhase -eq 'outlook-close') {
            $code = 'OUTLOOK_CLOSE_FAILED'
            $exitCode = 5
            $message = 'The owned Outlook session could not be closed normally. It was not forced closed; follow recovery instructions.'
        }
        elseif ($failurePhase -match '^(path-validate|output-claim|file-validate|publish|config-save|state-create|log-initialise)$') {
            $code = 'STORAGE_FAILED'
            $exitCode = 5
            $message = 'A file, path, permission or publication operation failed. The diagnostic report identifies the phase and error code.'
        }
        if ($hresult -eq '0x80070005' -or $hresult -eq '0x0000011F') {
            $code = 'ACCESS_DENIED'
            $exitCode = 3
            $message = 'Access was denied. The helper will not change policy or bypass the restriction.'
        }
        if ($types -contains 'System.Management.Automation.PipelineStoppedException') {
            $code = 'CANCELLED'
            $exitCode = 8
            $message = 'The operation was cancelled.'
        }
    }
    $line = 0
    if ($null -ne $Record.InvocationInfo) {
        $line = [int]$Record.InvocationInfo.ScriptLineNumber
    }
    return @{
        code           = $code
        message        = $message
        exitCode       = $exitCode
        hresult        = $hresult
        exceptionTypes = @($types.ToArray())
        line           = $line
        phase          = $failurePhase
    }
}

function Write-AtomicJson {
    param([string]$Path, $Value)
    $temporary = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    try {
        $text = ConvertTo-Json -InputObject $Value -Depth 12 -Compress
        [IO.File]::WriteAllText($temporary, $text, [Text.UTF8Encoding]::new($false))
        if ([IO.File]::Exists($Path)) {
            if (([IO.File]::GetAttributes($Path) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                Stop-Outcalget -Code 'UNSAFE_PATH' -Message 'A file is a reparse point; choose a regular local path.' -ExitCode 5
            }
            [IO.File]::Replace($temporary, $Path, [System.Management.Automation.Language.NullString]::Value)
        }
        else {
            [IO.File]::Move($temporary, $Path)
        }
    }
    finally {
        if ([IO.File]::Exists($temporary)) {
            try { [IO.File]::Delete($temporary) }
            catch {
                # Preserve the primary write/publication error; report cleanup at the boundary.
                $script:AtomicCleanupFailed = $true
            }
        }
    }
}

function Read-BoundedText {
    param([string]$Path, [long]$MaxBytes = 65536)
    $stream = $null
    $reader = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        if ($stream.Length -gt $MaxBytes) {
            Stop-Outcalget -Code 'INPUT_TOO_LARGE' -Message 'An input file exceeds its allowed size.' -ExitCode 2
        }
        $reader = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false, $true), $false)
        $text = $reader.ReadToEnd()
        return $text.TrimStart([char]0xFEFF)
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        elseif ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Read-FlatJson {
    # Local complexity exception: preserve the tested flat-schema duplicate-key validation; native JSON still decodes values.
    param([string]$Text)
    # Tokenise the deliberately flat scalar schema before using the built-in JSON decoder.
    # This detects duplicate/escaped duplicate keys that ConvertFrom-Json can otherwise discard.
    $stringPattern = '"(?:[^"\\\x00-\x1F]|\\(?:["\\/bfnrt]|u[0-9a-fA-F]{4}))*"'
    $tokenPattern = '\G[ \t\r\n]*(?<token>' + $stringPattern + '|true|false|null|-?(?:0|[1-9][0-9]*))'
    $tokenRegex = [regex]::new($tokenPattern, [Text.RegularExpressions.RegexOptions]::CultureInvariant, [TimeSpan]::FromSeconds(2))
    $values = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $position = 0
    $opening = [regex]::Match($Text, '\A[ \t\r\n]*\{')
    if (-not $opening.Success) {
        Stop-Outcalget -Code 'CONFIG_JSON' -Message 'Configuration must be one flat JSON object.' -ExitCode 2
    }
    $position = $opening.Length
    while ($true) {
        $tail = [regex]::Match($Text.Substring($position), '\A[ \t\r\n]*\}[ \t\r\n]*\z')
        if ($tail.Success) {
            break
        }
        $keyMatch = $tokenRegex.Match($Text, $position)
        if (-not $keyMatch.Success -or -not $keyMatch.Groups['token'].Value.StartsWith('"')) {
            Stop-Outcalget -Code 'CONFIG_JSON' -Message 'Configuration has invalid JSON property syntax.' -ExitCode 2
        }
        $keyToken = $keyMatch.Groups['token'].Value
        $key = (ConvertFrom-Json -InputObject ('{"v":' + $keyToken + '}')).v
        if (-not $seen.Add($key)) {
            Stop-Outcalget -Code 'CONFIG_DUPLICATE_KEY' -Message 'Configuration contains duplicate or case-colliding keys.' -ExitCode 2
        }
        $position = $keyMatch.Index + $keyMatch.Length
        $colon = [regex]::Match($Text.Substring($position), '\A[ \t\r\n]*:')
        if (-not $colon.Success) {
            Stop-Outcalget -Code 'CONFIG_JSON' -Message 'Configuration has invalid JSON property syntax.' -ExitCode 2
        }
        $position += $colon.Length
        $valueMatch = $tokenRegex.Match($Text, $position)
        if (-not $valueMatch.Success) {
            Stop-Outcalget -Code 'CONFIG_JSON' -Message 'Configuration values must be scalar strings, booleans or integers.' -ExitCode 2
        }
        $valueToken = $valueMatch.Groups['token'].Value
        $values.Add($key, (ConvertFrom-Json -InputObject ('{"v":' + $valueToken + '}')).v)
        $position = $valueMatch.Index + $valueMatch.Length
        $separator = [regex]::Match($Text.Substring($position), '\A[ \t\r\n]*(?<separator>[,}])')
        if (-not $separator.Success) {
            Stop-Outcalget -Code 'CONFIG_JSON' -Message 'Configuration contains invalid or non-integer JSON values.' -ExitCode 2
        }
        if ($separator.Groups['separator'].Value -eq '}') {
            if ($Text.Substring($position + $separator.Length) -notmatch '\A[ \t\r\n]*\z') {
                Stop-Outcalget -Code 'CONFIG_JSON' -Message 'Configuration has trailing data.' -ExitCode 2
            }
            break
        }
        $position += $separator.Length
        if ($Text.Substring($position) -match '\A[ \t\r\n]*\}') {
            Stop-Outcalget -Code 'CONFIG_JSON' -Message 'Trailing commas are not allowed.' -ExitCode 2
        }
        if ($values.Count -gt 32) {
            Stop-Outcalget -Code 'CONFIG_JSON' -Message 'Configuration has too many properties.' -ExitCode 2
        }
    }
    return , $values
}

function Test-Configuration {
    # Local complexity exception: exact conditional schemas preserve legacy bindings and fail closed on unknown source types.
    param($Values)
    $keys = @('schemaVersion', 'bindingId', 'userSid', 'profileName', 'outputPath',
        'includePrivateDetails', 'includeAttachments', 'timeoutSeconds', 'maxFileMiB')
    $strings = @('bindingId', 'userSid', 'profileName', 'outputPath')
    if ($Values['schemaVersion'] -eq 1) {
        $keys += @('storeId', 'folderEntryId')
        $strings += @('storeId', 'folderEntryId')
    }
    elseif ($Values['schemaVersion'] -eq 2) {
        $keys += 'sourceType'
        $strings += 'sourceType'
        switch -CaseSensitive ($Values['sourceType']) {
            'folder' {
                $keys += @('storeId', 'folderEntryId')
                $strings += @('storeId', 'folderEntryId')
            }
            'sharedDefaultCalendar' {
                $keys += @('sharedCalendarOwner', 'sharedOwnerEntryId')
                $strings += @('sharedCalendarOwner', 'sharedOwnerEntryId')
            }
            default {
                Stop-Outcalget -Code 'CONFIG_SOURCE' -Message 'The calendar source type is unsupported.' -ExitCode 2
            }
        }
    }
    else {
        Stop-Outcalget -Code 'CONFIG_SCHEMA' -Message 'Only configuration schemas 1 and 2 are supported.' -ExitCode 2
    }
    if ($Values.Count -ne $keys.Count) {
        Stop-Outcalget -Code 'CONFIG_SCHEMA' -Message 'Configuration properties do not match the selected schema and source type.' -ExitCode 2
    }
    foreach ($key in $keys) {
        if (-not $Values.ContainsKey($key)) {
            Stop-Outcalget -Code 'CONFIG_SCHEMA' -Message 'Configuration has missing, unknown or incorrectly cased properties.' -ExitCode 2
        }
    }
    foreach ($key in @('schemaVersion', 'timeoutSeconds', 'maxFileMiB')) {
        if ($Values[$key] -isnot [int] -and $Values[$key] -isnot [long]) {
            Stop-Outcalget -Code 'CONFIG_TYPE' -Message 'Configuration numeric values must be integers.' -ExitCode 2
        }
    }
    if ($Values['timeoutSeconds'] -lt 30 -or $Values['timeoutSeconds'] -gt 600 -or
        $Values['maxFileMiB'] -lt 1 -or $Values['maxFileMiB'] -gt 250) {
        Stop-Outcalget -Code 'CONFIG_RANGE' -Message 'Configuration version or numeric limits are unsupported.' -ExitCode 2
    }
    foreach ($key in @('includePrivateDetails', 'includeAttachments')) {
        if ($Values[$key] -isnot [bool]) {
            Stop-Outcalget -Code 'CONFIG_TYPE' -Message 'Privacy and attachment options must be JSON booleans.' -ExitCode 2
        }
    }
    foreach ($key in $strings) {
        if ($Values[$key] -isnot [string] -or [string]::IsNullOrWhiteSpace($Values[$key]) -or
            $Values[$key].Length -gt 16384 -or $Values[$key] -match '[\x00-\x1F<>]') {
            Stop-Outcalget -Code 'CONFIG_TYPE' -Message 'A required configuration string is invalid or contains a placeholder.' -ExitCode 2
        }
    }
    if ($Values['schemaVersion'] -eq 2 -and $Values['sourceType'] -ceq 'sharedDefaultCalendar') {
        Test-CalendarOwner -Owner $Values['sharedCalendarOwner']
    }
    $guid = [Guid]::Empty
    if (-not [Guid]::TryParseExact($Values['bindingId'], 'D', [ref]$guid) -or $guid -eq [Guid]::Empty -or
        $Values['userSid'] -notmatch '^S-1-[0-9]+(?:-[0-9]+)+$') {
        Stop-Outcalget -Code 'CONFIG_IDENTITY' -Message 'Configuration identity fields are invalid.' -ExitCode 2
    }
}

function Test-CalendarOwner {
    param([string]$Owner)
    try {
        $address = [Net.Mail.MailAddress]::new($Owner)
        if ($address.Address -ine $Owner -or $Owner.Length -gt 254 -or $Owner -match '[\s<>;,\x00-\x1F]' -or
            $Owner.IndexOf('@') -lt 1) {
            throw [FormatException]::new()
        }
    }
    catch {
        Stop-Outcalget -Code 'SHARED_OWNER_INVALID' -Message 'Enter one full calendar-owner SMTP address without a display name.' -ExitCode 2
    }
}

function Get-ExportConfigurationPath {
    param([string]$Name, [string]$ExplicitPath, [bool]$NameSpecified)
    if ($NameSpecified -and -not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        Stop-Outcalget -Code 'ARGUMENTS_INVALID' -Message 'Use -Profile or -Config, not both.' -ExitCode 2
    }
    if ($NameSpecified) {
        if ($Name -notmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,47}$' -or
            $Name -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') {
            Stop-Outcalget -Code 'EXPORT_PROFILE_INVALID' -Message 'Profile names need 1-48 ASCII letters, digits, hyphens or underscores, beginning with a letter or digit; reserved device names are forbidden.' -ExitCode 2
        }
        $Name = $Name.ToLowerInvariant()
        if ($Name -ne 'default') {
            $script:NamedExportProfile = $true
            return Join-Path -Path (Join-Path -Path $script:StateRoot -ChildPath 'profiles') -ChildPath ($Name + '.json')
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) { return $ExplicitPath }
    return Join-Path -Path $script:StateRoot -ChildPath 'config.json'
}

function Get-DateWindow {
    param([datetime]$Now)
    return @{
        start = $Now.Date
        end   = $Now.Date.AddDays(29)
    }
}

function ConvertTo-NativeArgument {
    param([string]$Value)
    $builder = [Text.StringBuilder]::new()
    [void]$builder.Append('"')
    $slashes = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq '\') {
            $slashes++
            continue
        }
        if ($character -eq '"') {
            [void]$builder.Append(('\' * (2 * $slashes + 1)))
            [void]$builder.Append('"')
        }
        else {
            if ($slashes -gt 0) {
                [void]$builder.Append(('\' * $slashes))
            }
            [void]$builder.Append($character)
        }
        $slashes = 0
    }
    if ($slashes -gt 0) {
        [void]$builder.Append(('\' * (2 * $slashes)))
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Resolve-LocalPath {
    param([string]$Path)
    if ($Path -notmatch '^[A-Za-z]:\\' -or $Path.Substring(2) -match '[:*?"<>|]' -or $Path -match '[\x00-\x1F]' -or $Path.Length -gt 230) {
        Stop-Outcalget -Code 'UNSAFE_PATH' -Message 'Use an absolute, regular local Windows path shorter than 231 characters.' -ExitCode 5
    }
    foreach ($part in $Path.Substring(3).Split('\')) {
        if ($part -match '[. ]$' -or $part -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)' -or $part.Contains('~')) {
            Stop-Outcalget -Code 'UNSAFE_PATH' -Message 'The path contains a reserved, ambiguous or short-name component.' -ExitCode 5
        }
    }
    $full = [IO.Path]::GetFullPath($Path)
    $drive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($full))
    if ($drive.DriveType -ne [IO.DriveType]::Fixed -or $drive.DriveFormat -ne 'NTFS') {
        Stop-Outcalget -Code 'FILESYSTEM_UNSUPPORTED' -Message 'Use a fixed local NTFS destination.' -ExitCode 5
    }
    $current = $full
    while (-not [string]::IsNullOrEmpty($current)) {
        if ([IO.File]::Exists($current) -or [IO.Directory]::Exists($current)) {
            if (([IO.File]::GetAttributes($current) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                Stop-Outcalget -Code 'REPARSE_POINT' -Message 'A path component is redirected or is a reparse point. Choose a regular local location.' -ExitCode 5
            }
        }
        $current = [IO.Path]::GetDirectoryName($current)
    }
    return $full
}

function Test-PrivateEntry {
    param([string]$Path)
    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
    $allowed = @($script:UserSid, 'S-1-5-18', 'S-1-5-32-544', 'S-1-3-0', 'S-1-3-4')
    $raw = [Security.AccessControl.RawSecurityDescriptor]::new($acl.GetSecurityDescriptorBinaryForm(), 0)
    if ($null -eq $raw.DiscretionaryAcl) {
        Stop-Outcalget -Code 'ACL_UNSAFE' -Message 'This location has no restrictive access-control list. Choose an approved private location.' -ExitCode 5
    }
    if ($acl.GetOwner([Security.Principal.SecurityIdentifier]).Value -notin @($script:UserSid, 'S-1-5-18', 'S-1-5-32-544')) {
        Stop-Outcalget -Code 'ACL_OWNER_UNSUPPORTED' -Message 'This location has an unexpected owner. Choose a private location owned by this user or the system administrators.' -ExitCode 5
    }
    $rules = $acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier])
    # Conservative: do not guess whether an unfamiliar corporate group is privileged.
    foreach ($rule in $rules) {
        if ($rule.AccessControlType -eq [Security.AccessControl.AccessControlType]::Allow -and
            ($rule.PropagationFlags -band [Security.AccessControl.PropagationFlags]::InheritOnly) -eq 0 -and
            ($allowed -notcontains $rule.IdentityReference.Value)) {
            Stop-Outcalget -Code 'ACL_REVIEW_REQUIRED' -Message 'This location grants access to an additional principal. Choose an approved private location; the helper will not rewrite existing access rules.' -ExitCode 5
        }
    }
}

function New-PrivateDirectory {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper under parent ShouldProcess; operational log directories are a documented preview side effect.')]
    [CmdletBinding()]
    param([string]$Path)
    $full = Resolve-LocalPath -Path $Path
    [void][IO.Directory]::CreateDirectory($full)
    Test-PrivateEntry -Path $full
    return $full
}

function Get-PathLock {
    param([string]$Path, [string]$Kind)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Path.ToUpperInvariant()))).Replace('-', '')
    }
    finally {
        $sha.Dispose()
    }
    $lockPath = Join-Path -Path $script:LockRoot -ChildPath ($Kind + '-' + $hash + '.lock')
    [void](Resolve-LocalPath -Path $lockPath)
    if ([IO.File]::Exists($lockPath)) {
        Test-PrivateEntry -Path $lockPath
    }
    try {
        return [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    }
    catch [IO.IOException] {
        Stop-Outcalget -Code 'RUN_BUSY' -Message 'Another operation holds this configuration or output lock. Try again after it finishes.' -ExitCode 6
    }
}

function Add-Trace {
    param([string]$Phase)
    $script:Phase = $Phase
    $script:Trace.Add(@{
            phase               = $Phase
            elapsedMilliseconds = $script:Clock.ElapsedMilliseconds
        })
    if ($script:Trace.Count -gt 128) {
        $script:Trace.RemoveAt(0)
    }
    if ($script:IsWorker -and $null -ne $script:ProgressPath) {
        Write-AtomicJson -Path $script:ProgressPath -Value @{
            runId  = $script:RunId
            phase  = $Phase
            events = @($script:Trace.ToArray())
            comAcquiredUtc = $script:ComAcquiredUtc
        }
    }
    elseif ($null -ne $script:RunLogPath) {
        Write-RunRecord -Completed $false
    }
    Write-PhaseMessage -Phase $Phase
}

function Wait-PathLock {
    param([string]$Path, [string]$Kind, [int]$TimeoutSeconds)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        try { return Get-PathLock -Path $Path -Kind $Kind }
        catch {
            if ($_.Exception.Data['OutcalgetCode'] -cne 'RUN_BUSY') { throw }
            if ($timer.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
                Stop-Outcalget -Code 'COORDINATION_TIMEOUT' -Message 'Timed out waiting for another export to release Outlook coordination. Try again after it finishes.' -ExitCode 7
            }
        }
        Start-Sleep -Milliseconds 100
    }
}

function Initialize-PathIdentityBridge {
    if ($null -ne ('Outcalget.PathIdentity' -as [type])) { return }
    # Read-only metadata: volume-GUID paths detect drive aliases; link count rejects hard-linked targets.
    # OPEN_EXISTING + zero access never creates/truncates a file. No privilege or policy changes.
    $source = @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace Outcalget {
    public sealed class PathMetadata {
        public string FinalPath;
        public uint Attributes;
        public uint Links;
    }
    public static class PathIdentity {
        [StructLayout(LayoutKind.Sequential)]
        private struct FileInformation {
            public uint Attributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME Creation;
            public System.Runtime.InteropServices.ComTypes.FILETIME Access;
            public System.Runtime.InteropServices.ComTypes.FILETIME Write;
            public uint VolumeSerial, SizeHigh, SizeLow, Links, IndexHigh, IndexLow;
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(string name, uint access, uint share,
            IntPtr security, uint disposition, uint flags, IntPtr template);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFinalPathNameByHandleW(SafeFileHandle handle,
            StringBuilder path, uint length, uint flags);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle handle,
            out FileInformation information);
        public static PathMetadata Read(string path) {
            // BACKUP_SEMANTICS permits directory handles; OPEN_REPARSE_POINT exposes links for rejection.
            using (SafeFileHandle handle = CreateFileW(path, 0, 7, IntPtr.Zero, 3,
                    0x02200000, IntPtr.Zero)) {
                if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
                FileInformation information;
                if (!GetFileInformationByHandle(handle, out information))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                StringBuilder buffer = new StringBuilder(1024);
                uint count = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 1);
                if (count == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
                if (count >= buffer.Capacity) throw new InvalidOperationException("Metadata path exceeds bound.");
                return new PathMetadata { FinalPath = buffer.ToString(),
                    Attributes = information.Attributes, Links = information.Links };
            }
        }
    }
}
'@
    try { Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop }
    catch {
        Stop-Outcalget -Code 'PATH_IDENTITY_UNAVAILABLE' -Message 'Existing policy prevented the read-only native path-identity adapter. No workaround was attempted.' -ExitCode 3
    }
}

function Get-CanonicalLocalPath {
    param([string]$Path)
    $full = Resolve-LocalPath -Path $Path
    Initialize-PathIdentityBridge
    $existing = $full
    $suffix = [Collections.Generic.List[string]]::new()
    while (-not [IO.File]::Exists($existing) -and -not [IO.Directory]::Exists($existing)) {
        $suffix.Insert(0, [IO.Path]::GetFileName($existing))
        $existing = [IO.Path]::GetDirectoryName($existing)
        if ([string]::IsNullOrEmpty($existing)) {
            Stop-Outcalget -Code 'PATH_IDENTITY_UNAVAILABLE' -Message 'The destination has no accessible local ancestor.' -ExitCode 5
        }
    }
    try { $metadata = [Outcalget.PathIdentity]::Read($existing) }
    catch {
        Stop-Outcalget -Code 'PATH_IDENTITY_UNAVAILABLE' -Message 'The destination path identity could not be verified.' -ExitCode 5
    }
    if (($metadata.Attributes -band 1024) -ne 0 -or (($metadata.Attributes -band 16) -eq 0 -and $metadata.Links -ne 1)) {
        Stop-Outcalget -Code 'PATH_ALIAS_UNSUPPORTED' -Message 'Reparse points and hard-linked configuration/output targets are not supported.' -ExitCode 5
    }
    $identity = $metadata.FinalPath.TrimEnd('\')
    foreach ($component in $suffix) { $identity += '\' + $component }
    return $identity.ToUpperInvariant()
}

function Get-ManagedConfiguration {
    $paths = [Collections.Generic.List[string]]::new()
    $default = Join-Path -Path $script:StateRoot -ChildPath 'config.json'
    if ([IO.File]::Exists($default)) { $paths.Add($default) }
    $profiles = Join-Path -Path $script:StateRoot -ChildPath 'profiles'
    if ([IO.Directory]::Exists($profiles)) {
        Test-PrivateEntry -Path $profiles
        $found = @(Get-ChildItem -LiteralPath $profiles -Filter '*.json' -File -ErrorAction Stop)
        if ($found.Count -gt 256) {
            Stop-Outcalget -Code 'PROFILE_LIMIT' -Message 'The profile catalog exceeds 256 configurations. Review it before continuing.' -ExitCode 2
        }
        foreach ($file in $found) { $paths.Add($file.FullName) }
    }
    return , $paths.ToArray()
}

function Test-OutputCatalog {
    param([string]$ConfigurationIdentity, [string]$OutputIdentity)
    foreach ($path in (Get-ManagedConfiguration)) {
        if ((Get-CanonicalLocalPath -Path $path) -ceq $ConfigurationIdentity) { continue }
        Test-PrivateEntry -Path $path
        $other = Read-FlatJson -Text (Read-BoundedText -Path $path -MaxBytes 65536)
        Test-Configuration -Values $other
        if ((Get-CanonicalLocalPath -Path $other['outputPath']) -ceq $OutputIdentity) {
            Stop-Outcalget -Code 'OUTPUT_IN_USE' -Message 'Another export configuration owns this output filename. Choose a separate destination for each profile.' -ExitCode 2
        }
    }
}

function Register-OutputClaim {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal reservation within the parent-approved export transaction; Check uses the read-only branch.')]
    [CmdletBinding()]
    param([string]$ConfigurationPath, [string]$OutputPath, [bool]$ReadOnly)
    $catalogLock = Wait-PathLock -Path (Join-Path -Path $script:StateRoot -ChildPath 'output-catalog') -Kind 'catalog' -TimeoutSeconds $CoordinationTimeoutSeconds
    try {
        $owner = Get-CanonicalLocalPath -Path $ConfigurationPath
        $identity = Get-CanonicalLocalPath -Path $OutputPath
        Test-OutputCatalog -ConfigurationIdentity $owner -OutputIdentity $identity
        $key = Get-CalendarFingerprint -StoreId 'output' -FolderId $identity
        $claimRoot = New-PrivateDirectory -Path (Join-Path -Path $script:StateRoot -ChildPath 'output-claims')
        $claimPath = Join-Path -Path $claimRoot -ChildPath ($key + '.json')
        $created = $false
        if ([IO.File]::Exists($claimPath)) {
            Test-PrivateEntry -Path $claimPath
            $claim = Read-FlatJson -Text (Read-BoundedText -Path $claimPath -MaxBytes 16384)
            $keys = @('schemaVersion', 'userSid', 'configurationIdentity', 'outputIdentity')
            if ($claim.Count -ne 4 -or @($keys | Where-Object { -not $claim.ContainsKey($_) }).Count -gt 0 -or
                $claim['schemaVersion'] -ne 1 -or $claim['userSid'] -cne $script:UserSid -or
                $claim['outputIdentity'] -cne $identity) {
                Stop-Outcalget -Code 'OUTPUT_CLAIM_INVALID' -Message 'An output ownership record is invalid. Review protected output-claims before retrying.' -ExitCode 5
            }
            if ($claim['configurationIdentity'] -cne $owner) {
                Stop-Outcalget -Code 'OUTPUT_IN_USE' -Message 'This filename is reserved by another export configuration, including configurations removed from the catalog. Choose another output filename.' -ExitCode 2
            }
        }
        elseif (-not $ReadOnly) {
            Write-AtomicJson -Path $claimPath -Value @{
                schemaVersion = 1
                userSid = $script:UserSid
                configurationIdentity = $owner
                outputIdentity = $identity
            }
            $created = $true
        }
        return [pscustomobject]@{ Path = $claimPath; Created = $created; Identity = $identity }
    }
    finally { $catalogLock.Dispose() }
}

function Invoke-ComRead {
    param([scriptblock]$Action)
    $delays = @(250, 500, 1000)
    for ($attempt = 0; ; $attempt++) {
        try {
            return (& $Action)
        }
        catch {
            $details = Get-ExceptionInfo -Record $_
            if ($details.hresult -notin @('0x80010001', '0x8001010A') -or $attempt -ge $delays.Count) {
                throw
            }
            Start-Sleep -Milliseconds $delays[$attempt]
        }
    }
}

function Test-IcsLine {
    # Local complexity exception: keep component state validation in one place to preserve native output acceptance.
    param([string]$Line, [hashtable]$State)
    if ($Line.Length -eq 0 -and $State.finished) {
        return
    }
    $match = [regex]::Match($Line, '^(?<name>[A-Za-z0-9-]+)(?<params>;[^:]*)?:(?<value>.*)$')
    if (-not $match.Success -or $State.finished) {
        Stop-Outcalget -Code 'ICS_INVALID' -Message 'The native export has invalid calendar framing or property syntax.' -ExitCode 5
    }
    $name = $match.Groups['name'].Value.ToUpperInvariant()
    $value = $match.Groups['value'].Value
    if ($name -eq 'BEGIN') {
        if ($value -notmatch '^[A-Za-z0-9-]+$' -or $match.Groups['params'].Success) {
            Stop-Outcalget -Code 'ICS_INVALID' -Message 'A calendar component declaration is invalid.' -ExitCode 5
        }
        $component = $value.ToUpperInvariant()
        if ($State.stack.Count -eq 0) {
            if ($State.started -or $component -ne 'VCALENDAR') {
                Stop-Outcalget -Code 'ICS_INVALID' -Message 'Expected a single VCALENDAR object.' -ExitCode 5
            }
            $State.started = $true
        }
        elseif ($component -eq 'VCALENDAR') {
            Stop-Outcalget -Code 'ICS_INVALID' -Message 'Nested calendar objects are not supported.' -ExitCode 5
        }
        if ($component -eq 'VEVENT') {
            $State.events++
        }
        if ($State.stack.Count -ge 32) {
            Stop-Outcalget -Code 'ICS_INVALID' -Message 'Calendar component nesting exceeds the allowed limit.' -ExitCode 5
        }
        $State.stack.Push($component)
    }
    elseif ($name -eq 'END') {
        if ($State.stack.Count -eq 0 -or $match.Groups['params'].Success -or $State.stack.Pop() -cne $value.ToUpperInvariant()) {
            Stop-Outcalget -Code 'ICS_INVALID' -Message 'Calendar component endings do not match.' -ExitCode 5
        }
        if ($State.stack.Count -eq 0) {
            $State.finished = $true
        }
    }
    else {
        if ($State.stack.Count -eq 0) {
            Stop-Outcalget -Code 'ICS_INVALID' -Message 'Calendar properties appear outside a component.' -ExitCode 5
        }
        if ($State.stack.Count -eq 1) {
            if ($name -eq 'VERSION') {
                if ($value -cne '2.0') {
                    Stop-Outcalget -Code 'ICS_VERSION' -Message 'Expected iCalendar version 2.0.' -ExitCode 5
                }
                $State.versions++
            }
            if ($name -eq 'PRODID' -and -not [string]::IsNullOrWhiteSpace($value)) {
                $State.products++
            }
        }
    }
}

function Test-IcsFile {
    # Local size/complexity exception: bounded reader, unfolding and disposal belong to one validation lifetime.
    param([string]$Path, [int]$MaxFileMiB)
    $stream = $null
    $reader = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $bytes = $stream.Length
        if ($bytes -eq 0 -or $bytes -gt ([long]$MaxFileMiB * 1048576)) {
            Stop-Outcalget -Code 'ICS_SIZE' -Message 'The native export is empty or exceeds the configured size limit.' -ExitCode 5
        }
        $reader = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false, $true), $true)
        $state = @{
            stack    = [Collections.Stack]::new()
            started  = $false
            finished = $false
            versions = 0
            products = 0
            events   = 0
        }
        $pending = [Text.StringBuilder]::new()
        $hasPending = $false
        $lines = 0
        while ($null -ne ($line = $reader.ReadLine())) {
            $lines++
            if ($lines -gt 1000000 -or $line.Length -gt 1048576) {
                Stop-Outcalget -Code 'ICS_LIMIT' -Message 'The calendar exceeds a structural validation limit.' -ExitCode 5
            }
            if ($line.StartsWith(' ') -or $line.StartsWith("`t")) {
                if (-not $hasPending) {
                    Stop-Outcalget -Code 'ICS_INVALID' -Message 'An iCalendar continuation line has no preceding property.' -ExitCode 5
                }
                [void]$pending.Append($line.Substring(1))
            }
            else {
                if ($hasPending) {
                    Test-IcsLine -Line $pending.ToString() -State $state
                }
                [void]$pending.Clear()
                [void]$pending.Append($line)
                $hasPending = $true
            }
            if ($pending.Length -gt 1048576) {
                Stop-Outcalget -Code 'ICS_LIMIT' -Message 'An unfolded calendar line exceeds 1 MiB.' -ExitCode 5
            }
        }
        if ($hasPending) {
            Test-IcsLine -Line $pending.ToString() -State $state
        }
        if (-not $state.finished -or $state.stack.Count -ne 0 -or $state.versions -ne 1 -or $state.products -ne 1) {
            Stop-Outcalget -Code 'ICS_INVALID' -Message 'The native export is incomplete or lacks required calendar properties.' -ExitCode 5
        }
        return @{
            bytes               = $bytes
            eventComponentCount = $state.events
        }
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        elseif ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Publish-CalendarFile {
    param([string]$Stage, [string]$Destination, [bool]$MustBeNew)
    if ([IO.Directory]::Exists($Destination)) {
        Stop-Outcalget -Code 'DESTINATION_IS_DIRECTORY' -Message 'The configured destination is a directory.' -ExitCode 5
    }
    if ([IO.File]::Exists($Destination)) {
        if ($MustBeNew) {
            Stop-Outcalget -Code 'DESTINATION_EXISTS' -Message 'Setup requires an unoccupied output filename.' -ExitCode 5
        }
        [IO.File]::Replace($Stage, $Destination, [System.Management.Automation.Language.NullString]::Value)
    }
    else {
        [IO.File]::Move($Stage, $Destination)
    }
}

function Test-HostEnvironment {
    if ($env:OS -ne 'Windows_NT' -or $PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1) {
        Stop-Outcalget -Code 'HOST_UNSUPPORTED' -Message 'Use permitted Windows PowerShell 5.1 (powershell.exe) on Windows.' -ExitCode 3
    }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = [Security.Principal.WindowsPrincipal]::new($identity)
        $script:SessionInteractive = $principal.IsInRole([Security.Principal.SecurityIdentifier]::new('S-1-5-4'))
        $script:SessionElevated = $identity.IsSystem -or $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $script:UserSid = $identity.User.Value
    }
    finally {
        $identity.Dispose()
    }
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
        Stop-Outcalget -Code 'STA_REQUIRED' -Message 'Start Windows PowerShell with -STA.' -ExitCode 3
    }
}

function Test-ExportApproval {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param()
    # A bound -Debug can force ShouldProcess confirmation independently of DebugPreference.
    # This adapter is called with -Debug:$false; explicit Confirm/WhatIf are forwarded.
    return $PSCmdlet.ShouldProcess('configured calendar and local export', 'Export with Outlook Classic; save configuration only with -Configure')
}

function Test-DesktopSession {
    # Interactive-token tasks share the user's desktop. Password/S4U tasks do not.
    if ($script:SessionElevated) {
        Stop-Outcalget -Code 'ELEVATION_UNSUPPORTED' -Message 'Use the configured ordinary user. Clear Run with highest privileges in Task Scheduler.' -ExitCode 3
    }
    if (-not $script:SessionInteractive -or -not [Environment]::UserInteractive -or [Diagnostics.Process]::GetCurrentProcess().SessionId -eq 0) {
        Stop-Outcalget -Code 'SESSION_UNSUPPORTED' -Message 'Select Run only when user is logged on, under the same user as the configured running Outlook Classic session.' -ExitCode 3
    }
}

function Write-PhaseMessage {
    param([string]$Phase)
    if ($script:IsWorker -or -not $script:ConsoleFeedback) {
        return
    }
    $messages = @{
        'host-check'          = 'Checking Windows PowerShell.'
        'session-check'       = 'Checking the signed-in desktop session.'
        'config-load'         = 'Loading the saved configuration.'
        'internet-check'      = 'Checking Windows Internet connectivity.'
        'output-claim'        = 'Checking exclusive ownership of the output filename.'
        'coordination-wait'   = 'Waiting for Outlook lifecycle coordination.'
        'com-wait'            = 'Waiting for the Outlook COM operation queue.'
        'outlook-joined'      = 'Joining the Outlook session started by another export; keeping it open until all participants finish.'
        'outlook-users-wait'  = 'Finishing Outlook cleanup; waiting for joined exports when required.'
        'outlook-existing'    = 'Outlook Classic is already open; leaving its session unchanged.'
        'outlook-launch'      = 'Starting Outlook Classic.'
        'outlook-ready-wait'  = 'Waiting for Outlook Classic and the calendar to be ready.'
        'outlook-setup-wait'  = 'Waiting for the Outlook profile, stores and required connection before calendar selection.'
        'outlook-online-wait' = 'Waiting for Outlook to connect to Exchange.'
        'sync-start'          = 'Starting synchronization for the selected calendar.'
        'sync-prepare'        = 'Preparing the native synchronization event monitor.'
        'sync-wait'           = 'Waiting for synchronization completion.'
        'sync-connection-wait' = 'Outlook is not fully connected; waiting within the synchronization timeout.'
        'sync-connection-recovered' = 'Outlook connection recovered; requiring a new completed synchronization.'
        'sync-complete'       = 'Calendar synchronization group completed without a reported error.'
        'sync-restore'        = 'Restoring the calendar synchronization setting.'
        'ical-save'           = 'Exporting the calendar with Full details.'
        'file-validate'       = 'Checking the staged calendar file.'
        'publish'             = 'Publishing the complete calendar file.'
        'outlook-close'       = 'Closing the Outlook Classic session started by this run.'
        'outlook-closed'      = 'The owned Outlook Classic process has exited.'
        'complete'            = 'Requested export work completed; finishing cleanup and run logging.'
    }
    if ($messages.ContainsKey($Phase)) {
        try {
            $stamp = [datetime]::UtcNow.ToString('HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
            Write-Information -InformationAction Continue -MessageData ('[{0} UTC] {1}' -f $stamp, $messages[$Phase])
        }
        catch {
            # Console feedback is optional; the required sanitized run record remains authoritative.
            $script:ConsoleFeedbackFailed = $true
        }
    }
}

function Test-InternetConnection {
    # Local complexity exception: fail-closed status/type/error classification and native-reference cleanup form one check.
    $network = $null
    try {
        $type = [Type]::GetTypeFromCLSID([Guid]::new('DCB00C01-570F-4A9B-8D69-199FDBA5723B'), $true)
        $network = [Activator]::CreateInstance($type)
        $available = $network.IsConnectedToInternet
        if ($available -isnot [bool]) {
            Stop-Outcalget -Code 'INTERNET_CHECK_FAILED' -Message 'Windows could not report Internet connectivity. No calendar export was attempted.' -ExitCode 3
        }
        if (-not $available) {
            Stop-Outcalget -Code 'INTERNET_UNAVAILABLE' -Message 'Windows reports no Internet connection. Connect normally and retry; offline export is disabled.' -ExitCode 3
        }
    }
    catch {
        $details = Get-ExceptionInfo -Record $_
        if ($details.code -in @('INTERNET_UNAVAILABLE', 'INTERNET_CHECK_FAILED')) {
            throw
        }
        Stop-Outcalget -Code 'INTERNET_CHECK_FAILED' -Message 'The Windows Internet connectivity check was unavailable or blocked. The helper will not bypass policy or export offline.' -ExitCode 3
    }
    finally {
        if ($null -ne $network) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($network) }
            catch { $script:ComCleanupWarning = $true }
        }
    }
}

function Get-OutlookInstance {
    $current = [Diagnostics.Process]::GetCurrentProcess()
    try { $sessionId = $current.SessionId }
    finally { $current.Dispose() }
    $instances = [Collections.Generic.List[object]]::new()
    foreach ($process in [Diagnostics.Process]::GetProcessesByName('OUTLOOK')) {
        try {
            if ($process.SessionId -eq $sessionId) {
                $instances.Add(@{
                        id           = $process.Id
                        startedTicks = $process.StartTime.ToUniversalTime().Ticks
                    })
            }
        }
        catch {
            Stop-Outcalget -Code 'OUTLOOK_PROCESS_UNAVAILABLE' -Message 'An Outlook process could not be identified safely. No additional Outlook instance will be started or closed.' -ExitCode 3
        }
        finally { $process.Dispose() }
    }
    return $instances.ToArray()
}

function Test-OutlookLease {
    param($Lease)
    if ($null -eq $Lease -or ($Lease.id -isnot [int] -and $Lease.id -isnot [long]) -or $Lease.id -le 0 -or $Lease.id -gt [int]::MaxValue -or
        ($Lease.startedTicks -isnot [long] -and $Lease.startedTicks -isnot [int]) -or $Lease.startedTicks -le 0) {
        Stop-Outcalget -Code 'OUTLOOK_OWNERSHIP_UNCERTAIN' -Message 'Outlook process ownership could not be verified. The helper will not close an uncertain session.' -ExitCode 3
    }
    $instances = @(Get-OutlookInstance)
    if ($instances.Count -ne 1 -or $instances[0].id -ne $Lease.id -or $instances[0].startedTicks -ne $Lease.startedTicks) {
        Stop-Outcalget -Code 'OUTLOOK_OWNERSHIP_CHANGED' -Message 'The Outlook session changed during this run. The helper will leave it open.' -ExitCode 3
    }
}

function Get-OutlookExecutable {
    # Only machine-controlled App Paths are used; never PATH, a shell association or HKCU.
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
        $root = $null
        $key = $null
        try {
            $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
            $key = $root.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE')
            if ($null -eq $key) { continue }
            $path = $key.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            if ($path -is [string]) {
                $path = [Environment]::ExpandEnvironmentVariables($path).Trim().Trim('"')
                if ([IO.Path]::IsPathRooted($path) -and [IO.Path]::GetFileName($path) -ieq 'OUTLOOK.EXE' -and [IO.File]::Exists($path)) {
                    return [IO.Path]::GetFullPath($path)
                }
            }
        }
        finally {
            if ($null -ne $key) { $key.Dispose() }
            if ($null -ne $root) { $root.Dispose() }
        }
    }
    Stop-Outcalget -Code 'OUTLOOK_INSTALLATION_UNAVAILABLE' -Message 'A machine-registered Outlook Classic executable was not found. Open the approved Outlook installation normally or ask IT to repair its registration.' -ExitCode 3
}

function Start-OutlookSession {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal lifecycle operation under the parent export ShouldProcess and exclusive Outlook lock.')]
    [CmdletBinding()]
    param([string]$ProfileName)
    $info = [Diagnostics.ProcessStartInfo]::new()
    $info.FileName = Get-OutlookExecutable
    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        # Interactive setup has no saved binding yet: use normal Outlook startup.
        $info.Arguments = ''
    }
    else {
        $info.Arguments = '/profile ' + (ConvertTo-NativeArgument -Value $ProfileName)
    }
    $info.UseShellExecute = $false
    $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Minimized
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $info
    try {
        if (@(Get-OutlookInstance).Count -ne 0) {
            Stop-Outcalget -Code 'OUTLOOK_START_RACE' -Message 'An Outlook session appeared during startup. It was left untouched; retry after startup finishes.' -ExitCode 3
        }
        $script:OutlookStartAttempted = $true
        if (-not $process.Start()) {
            Stop-Outcalget -Code 'OUTLOOK_START_FAILED' -Message 'Outlook Classic could not be started.' -ExitCode 3
        }
        $script:OwnedOutlookProcess = $process
        $script:RunResult.outlookStarted = $true
        $lease = @{ id = $process.Id; startedTicks = $process.StartTime.ToUniversalTime().Ticks }
        $script:OutlookLease = $lease
        return $lease
    }
    finally {
        if ($null -eq $script:OwnedOutlookProcess) { $process.Dispose() }
    }
}

function Get-ReadyOutlookApplication {
    param([int]$TimeoutSeconds = 30)
    Add-Trace -Phase 'outlook-ready-wait'
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        }
        catch {
            $details = Get-ExceptionInfo -Record $_
            if ($details.hresult -notin @('0x800401E3', '0x80010001', '0x8001010A')) { throw }
        }
        Start-Sleep -Milliseconds 250
    }
    Stop-Outcalget -Code 'OUTLOOK_STARTUP_TIMEOUT' -Message 'Outlook Classic did not become available in time. Finish any sign-in or profile dialog interactively before retrying.' -ExitCode 7
}

function Wait-OutlookConnection {
    param($Session, [int]$TimeoutSeconds)
    Add-Trace -Phase 'outlook-online-wait'
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        Test-InternetConnection
        $mode = Invoke-ComRead -Action { [int]$Session.ExchangeConnectionMode }
        if (-not $Session.Offline -and $mode -in @(600, 700, 800)) { return }
        Start-Sleep -Milliseconds 500
    }
    Stop-Outcalget -Code 'OUTLOOK_CONNECTION_TIMEOUT' -Message 'Outlook did not reach an online full-item Exchange connection in time. The cached calendar was not exported.' -ExitCode 7
}

function Get-ReadyOutlookSetup {
    # Local complexity exception: one bounded poll owns profile/store/connection checks and the per-attempt Stores reference lifetime.
    param($Session, [bool]$RequireOnline, [int]$TimeoutSeconds)
    Add-Trace -Phase 'outlook-setup-wait'
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $stores = $null
        try {
            $profileName = [string]$Session.CurrentProfileName
            $stores = $Session.Stores
            $storeCount = 0
            if ($null -ne $stores) { $storeCount = [int]$stores.Count }
            $online = -not $RequireOnline
            if ($RequireOnline -and $storeCount -gt 0 -and -not [string]::IsNullOrWhiteSpace($profileName)) {
                $online = -not $Session.Offline -and [int]$Session.ExchangeConnectionMode -in @(600, 700, 800)
            }
            if (-not [string]::IsNullOrWhiteSpace($profileName) -and $storeCount -gt 0 -and $online) {
                Test-InternetConnection
                return $profileName
            }
        }
        catch {
            $details = Get-ExceptionInfo -Record $_
            if ($details.hresult -notin @('0x80010001', '0x8001010A', '0x8004010F', '0x80040111', '0x80040115')) { throw }
        }
        finally { Close-ComReference -Reference $stores }
        Test-InternetConnection
        Start-Sleep -Milliseconds 250
    }
    Stop-Outcalget -Code 'OUTLOOK_STARTUP_TIMEOUT' -Message 'Outlook did not make its profile, stores or required Exchange connection available for setup in time. Complete any startup or sign-in dialog interactively and retry.' -ExitCode 7
}

function Get-ReadyOutlookSource {
    # A newly started profile can publish its COM application before stores load.
    param($Session, $Settings, [int]$TimeoutSeconds)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $outlookProfileName = Invoke-ComRead -Action { [string]$Session.CurrentProfileName }
        if ([string]::IsNullOrEmpty($outlookProfileName)) {
            Test-InternetConnection
            Start-Sleep -Milliseconds 250
            continue
        }
        if ($outlookProfileName -cne $Settings['profileName']) {
            Stop-Outcalget -Code 'PROFILE_MISMATCH' -Message 'The started Outlook profile does not match the saved configuration.' -ExitCode 2
        }
        try { return Resolve-ConfiguredCalendar -Session $Session -Settings $Settings }
        catch {
            $details = Get-ExceptionInfo -Record $_
            if ($details.hresult -notin @('0x80010001', '0x8001010A', '0x8004010F', '0x80040111', '0x80040115')) { throw }
        }
        Test-InternetConnection
        Start-Sleep -Milliseconds 250
    }
    Stop-Outcalget -Code 'OUTLOOK_STARTUP_TIMEOUT' -Message 'The saved Outlook profile or calendar store did not become available in time. Finish startup or sign-in interactively and check the saved binding.' -ExitCode 7
}

function Close-ComReference {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Releases only the COM reference acquired by the calling adapter.')]
    [CmdletBinding()]
    param($Reference)
    if ($null -ne $Reference -and [Runtime.InteropServices.Marshal]::IsComObject($Reference)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Reference) }
        catch { $script:ComCleanupWarning = $true }
    }
}

function Get-SharedCalendarBinding {
    # Local complexity exception: recipient resolution, mailbox identity pinning and transferred folder ownership share one COM-reference lifetime.
    param($Session, [string]$Owner, [string]$ExpectedOwnerId)
    Test-CalendarOwner -Owner $Owner
    $recipient = $null
    $addressEntry = $null
    $exchangeUser = $null
    $folder = $null
    try {
        $recipient = $Session.CreateRecipient($Owner)
        if (-not $recipient.Resolve() -or -not $recipient.Resolved) {
            Stop-Outcalget -Code 'SHARED_OWNER_UNRESOLVED' -Message 'Outlook could not resolve the calendar owner. Check the full email address and directory access.' -ExitCode 3
        }
        $addressEntry = $recipient.AddressEntry
        # This binding targets an Exchange directory mailbox, rather than an ambiguous display name or one-off recipient.
        if ([string]$addressEntry.Type -cne 'EX') {
            Stop-Outcalget -Code 'SHARED_OWNER_NOT_EXCHANGE' -Message 'The resolved owner is not an Exchange directory mailbox. Try selecting the already-added calendar from the Calendar navigation list.' -ExitCode 2
        }
        $exchangeUser = $addressEntry.GetExchangeUser()
        if ($null -eq $exchangeUser) {
            Stop-Outcalget -Code 'SHARED_OWNER_UNAVAILABLE' -Message 'The resolved Exchange calendar owner could not be inspected.' -ExitCode 3
        }
        $smtp = [string]$exchangeUser.PrimarySmtpAddress
        $ownerId = [string]$addressEntry.ID
        Test-CalendarOwner -Owner $smtp
        # Entry IDs can have different binary representations for the same object.
        # Ask MAPI to compare the pinned owner rather than compare the strings.
        $ownerMatched = [string]::IsNullOrEmpty($ExpectedOwnerId)
        if (-not $ownerMatched -and -not [string]::IsNullOrWhiteSpace($ownerId)) {
            $ownerMatched = Invoke-ComRead -Action { [bool]$Session.CompareEntryIDs($ownerId, $ExpectedOwnerId) }
        }
        if ([string]::IsNullOrWhiteSpace($ownerId) -or -not $ownerMatched) {
            Stop-Outcalget -Code 'SHARED_OWNER_CHANGED' -Message 'The resolved mailbox identity no longer matches the configured owner. Reconfigure explicitly after reviewing the change.' -ExitCode 2
        }
        try { $folder = $Session.GetSharedDefaultFolder($recipient, 9) }
        catch {
            Stop-Outcalget -Code 'SHARED_CALENDAR_UNAVAILABLE' -Message 'Outlook cannot open the shared default calendar. Confirm sharing permissions and test access in Outlook Classic.' -ExitCode 3 -InnerException $_.Exception
        }
        if ($null -eq $folder -or [int]$folder.DefaultItemType -ne 1) {
            Stop-Outcalget -Code 'NOT_A_CALENDAR' -Message 'The shared default folder is not an accessible calendar.' -ExitCode 2
        }
        $binding = [pscustomobject]@{ Folder = $folder; Owner = $smtp; OwnerEntryId = $ownerId }
        $folder = $null
        return $binding
    }
    finally {
            foreach ($reference in @($folder, $exchangeUser, $addressEntry, $recipient)) {
            Close-ComReference -Reference $reference
        }
    }
}

function Resolve-ConfiguredCalendar {
    param($Session, $Settings)
    if ($Settings['schemaVersion'] -eq 2 -and $Settings['sourceType'] -ceq 'sharedDefaultCalendar') {
        $binding = Get-SharedCalendarBinding -Session $Session -Owner $Settings['sharedCalendarOwner'] -ExpectedOwnerId $Settings['sharedOwnerEntryId']
        return $binding.Folder
    }
    $folder = $Session.GetFolderFromID($Settings['folderEntryId'], $Settings['storeId'])
    if ([string]$folder.StoreID -cne $Settings['storeId'] -or [string]$folder.EntryID -cne $Settings['folderEntryId']) {
        Close-ComReference -Reference $folder
        Stop-Outcalget -Code 'SOURCE_MISMATCH' -Message 'The resolved calendar does not match the saved folder binding.' -ExitCode 2
    }
    return $folder
}

function Test-CalendarExporterSource {
    [CmdletBinding()]
    param($Exporter, $Folder, $Settings, $Session)
    $exporterFolder = $null
    $nativeExporter = $Exporter
    $nativeSession = $Session
    try {
        $exporterFolder = Invoke-ComRead -Action { $nativeExporter.Folder }
        if ($null -eq $exporterFolder -or [int]$exporterFolder.DefaultItemType -ne 1) {
            Stop-Outcalget -Code 'EXPORT_SOURCE_UNAVAILABLE' -Message 'Outlook did not bind its calendar exporter to an accessible calendar folder.' -ExitCode 4
        }
        $expectedEntryId = [string]$Folder.EntryID
        $actualEntryId = [string]$exporterFolder.EntryID
        if ([string]::IsNullOrWhiteSpace($expectedEntryId) -or [string]::IsNullOrWhiteSpace($actualEntryId)) {
            Stop-Outcalget -Code 'EXPORT_SOURCE_UNAVAILABLE' -Message 'The calendar exporter source could not be identified before export.' -ExitCode 4
        }
        $sameFolder = Invoke-ComRead -Action { [bool]$nativeSession.CompareEntryIDs($actualEntryId, $expectedEntryId) }
        if (-not $sameFolder) {
            Stop-Outcalget -Code 'EXPORT_SOURCE_MISMATCH' -Message 'Outlook attached the calendar exporter to a different folder than the configured calendar.' -ExitCode 4
        }
        if (-not ($Settings['schemaVersion'] -eq 2 -and $Settings['sourceType'] -ceq 'sharedDefaultCalendar')) {
            $expectedStoreId = [string]$Folder.StoreID
            $actualStoreId = [string]$exporterFolder.StoreID
            if ([string]::IsNullOrWhiteSpace($expectedStoreId) -or
                $actualStoreId -cne $expectedStoreId -or $actualStoreId -cne [string]$Settings['storeId']) {
                Stop-Outcalget -Code 'EXPORT_SOURCE_MISMATCH' -Message 'Outlook attached the calendar exporter to a different store than the configured calendar.' -ExitCode 4
            }
        }
    }
    finally {
        Close-ComReference -Reference $exporterFolder
    }
}

function Get-CalendarFingerprint {
    param([string]$StoreId, [string]$FolderId)
    $bytes = [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -Compress -InputObject @($StoreId, $FolderId)))
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hash.ComputeHash($bytes)).Replace('-', '') }
    finally { $hash.Dispose() }
}

function Get-NavigationCalendar {
    # Local complexity exception: indexed COM traversal owns and releases each navigation reference in one scope.
    param($Application, [string]$Fingerprint)
    $explorer = $null
    $pane = $null
    $modules = $null
    $module = $null
    $groups = $null
    $entries = [Collections.Generic.List[object]]::new()
    $seen = @{}
    try {
        $explorer = $Application.ActiveExplorer()
        if ($null -eq $explorer) {
            Stop-Outcalget -Code 'NAVIGATION_UNAVAILABLE' -Message 'Open the Calendar view in Outlook Classic before choosing a calendar from its navigation list.' -ExitCode 3
        }
        $pane = $explorer.NavigationPane
        $modules = $pane.Modules
        $module = $modules.GetNavigationModule(1)
        $groups = $module.NavigationGroups
        for ($groupIndex = 1; $groupIndex -le [int]$groups.Count; $groupIndex++) {
            $group = $null
            $navigationFolders = $null
            try {
                $group = $groups.Item($groupIndex)
                $navigationFolders = $group.NavigationFolders
                for ($index = 1; $index -le [int]$navigationFolders.Count; $index++) {
                    $navigationFolder = $null
                    $folder = $null
                    try {
                        $navigationFolder = $navigationFolders.Item($index)
                        $folder = $navigationFolder.Folder
                        if ($null -eq $folder -or [int]$folder.DefaultItemType -ne 1) { continue }
                        $key = Get-CalendarFingerprint -StoreId ([string]$folder.StoreID) -FolderId ([string]$folder.EntryID)
                        if ($Fingerprint -and $key -ceq $Fingerprint) {
                            $selected = $folder
                            $folder = $null
                            return $selected
                        }
                        if (-not $seen.ContainsKey($key)) {
                            $seen[$key] = $true
                            $label = [string]$group.Name + ' / ' + [string]$navigationFolder.DisplayName
                            $label = $label -replace '[\x00-\x1F\x7F]', ' '
                            if ($label.Length -gt 160) { $label = $label.Substring(0, 160) }
                            $entries.Add([pscustomobject]@{ label = $label; fingerprint = $key })
                            if ($entries.Count -gt 200) {
                                Stop-Outcalget -Code 'NAVIGATION_LIMIT' -Message 'The calendar navigation list exceeds 200 distinct calendars. Use the folder picker or shared-default option.' -ExitCode 2
                            }
                        }
                    }
                    catch {
                        $failure = Get-ExceptionInfo -Record $_
                        if ($failure.hresult -notin @('0x8004010F', '0x80070005', '0x80040111')) { throw }
                        $script:NavigationUnavailableCount++
                    }
                    finally {
                        Close-ComReference -Reference $folder
                        Close-ComReference -Reference $navigationFolder
                    }
                }
            }
            finally {
                Close-ComReference -Reference $navigationFolders
                Close-ComReference -Reference $group
            }
        }
        if ($Fingerprint) {
            Stop-Outcalget -Code 'NAVIGATION_CHANGED' -Message 'The selected calendar is no longer in the navigation list. Run configuration again.' -ExitCode 2
        }
        return $entries.ToArray()
    }
    finally {
        foreach ($reference in @($groups, $module, $modules, $pane, $explorer)) {
            Close-ComReference -Reference $reference
        }
    }
}

function Initialize-SynchronizationBridge {
    if ($null -ne ('Outcalget.SyncMonitor' -as [type])) { return }
    # A narrow native-event bridge avoids an optional Outlook PIA dependency and
    # PowerShell callbacks on COM threads. GUID/DISPIDs verified against Microsoft's
    # Outlook interop metadata: SyncStart 61441, OnError 61443, SyncEnd 61444.
    $source = @'
using System;
using System.Runtime.InteropServices;
using System.Threading;
namespace Outcalget {
    public sealed class SyncMonitor : IDisposable {
        private static readonly Guid Events = new Guid("00063085-0000-0000-C000-000000000046");
        private readonly object source;
        private readonly Action onStart;
        private readonly Action onEnd;
        private readonly Action<int, string> onError;
        private int started, ended, errors, lastError, disposed;
        private bool startHook, errorHook, endHook;
        public bool Started { get { return Volatile.Read(ref started) != 0; } }
        public bool Completed { get { return Started && Volatile.Read(ref ended) != 0; } }
        public int ErrorCount { get { return Volatile.Read(ref errors); } }
        public int LastError { get { return Volatile.Read(ref lastError); } }
        public SyncMonitor(object value) {
            source = value;
            onStart = delegate { Interlocked.Exchange(ref ended, 0); Interlocked.Exchange(ref started, 1); };
            onEnd = delegate { if (Started) Interlocked.Exchange(ref ended, 1); };
            onError = delegate(int code, string description) {
                Interlocked.Exchange(ref lastError, code);
                Interlocked.Increment(ref errors);
            };
            try {
                ComEventsHelper.Combine(source, Events, 61441, onStart); startHook = true;
                ComEventsHelper.Combine(source, Events, 61443, onError); errorHook = true;
                ComEventsHelper.Combine(source, Events, 61444, onEnd); endHook = true;
            } catch {
                // Keep the original subscription failure if best-effort detach also fails;
                // the parent treats startup/sync as failed and attempts owned-session Quit.
                try { Dispose(); } catch { } throw;
            }
        }
        public void Dispose() {
            if (Interlocked.Exchange(ref disposed, 1) != 0) return;
            Exception failure = null;
            if (endHook) try { ComEventsHelper.Remove(source, Events, 61444, onEnd); } catch (Exception e) { failure = e; }
            if (errorHook) try { ComEventsHelper.Remove(source, Events, 61443, onError); } catch (Exception e) { if (failure == null) failure = e; }
            if (startHook) try { ComEventsHelper.Remove(source, Events, 61441, onStart); } catch (Exception e) { if (failure == null) failure = e; }
            if (failure != null) throw failure;
        }
    }
}
'@
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -TypeDefinition $source -ErrorAction Stop
    }
    catch {
        Stop-Outcalget -Code 'SYNC_BRIDGE_UNAVAILABLE' -Message 'The native synchronization event bridge could not be loaded. No policy workaround or timed-delay fallback will be used.' -ExitCode 3
    }
}

function Get-OutlookConnectionState {
    param($Session)
    $offline = [bool]$Session.Offline
    $mode = [int]$Session.ExchangeConnectionMode
    return [pscustomobject]@{
        offline = $offline
        mode = $mode
        ready = -not $offline -and $mode -in @(600, 700, 800)
    }
}

function Add-ConnectionObservation {
    param($State, $Observations, [long]$ElapsedMilliseconds, [int]$Attempt)
    if ($Observations.Count -gt 0) {
        $last = $Observations[$Observations.Count - 1]
        if ($last.offline -eq $State.offline -and $last.primaryExchangeConnectionMode -eq $State.mode -and $last.syncAttempt -eq $Attempt) { return }
    }
    if ($Observations.Count -ge 24) { $Observations.RemoveAt(0) }
    $Observations.Add(@{
        offline = $State.offline
        primaryExchangeConnectionMode = $State.mode
        elapsedMilliseconds = $ElapsedMilliseconds
        syncAttempt = $Attempt
    })
}

function Close-SynchronizationAttempt {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal stop/unsubscribe of the native synchronization attempt owned by the caller; part of existing parent-authorized lifecycle cleanup.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleTypes', '', Justification = 'Initialize-SynchronizationBridge explicitly loads the Windows .NET Framework System.Windows.Forms assembly before synchronization cleanup.')]
    [CmdletBinding()]
    param($SyncObject, $Monitor, [bool]$Started, [long]$TimeoutMilliseconds = 180000)
    $failure = $null
    try {
        if ($Started -and -not $Monitor.Completed) {
            $timer = [Diagnostics.Stopwatch]::StartNew()
            $SyncObject.Stop()
            while (-not $Monitor.Completed) {
                if ($timer.ElapsedMilliseconds -ge $TimeoutMilliseconds) {
                    throw [TimeoutException]::new('The interrupted synchronization did not finish within the remaining budget.')
                }
                [Windows.Forms.Application]::DoEvents()
                if (-not $Monitor.Completed) { Start-Sleep -Milliseconds 25 }
            }
        }
    }
    catch { $failure = $_.Exception }
    try { $Monitor.Dispose() }
    catch { if ($null -eq $failure) { $failure = $_.Exception } }
    if ($null -ne $failure) {
        $script:SyncCleanupFailed = $true
        Stop-Outcalget -Code 'SYNC_CLEANUP_FAILED' -Message 'The interrupted synchronization attempt could not be stopped or detached safely. No new attempt or export will start.' -ExitCode 5 -InnerException $failure
    }
}

function Wait-CalendarSynchronization {
    # Local complexity exception: bounded connection recovery, correlated events and cleanup share one synchronization lifetime and deadline.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleTypes', '', Justification = 'The Desktop 5.1 profile lists default-loaded types only. Initialize-SynchronizationBridge explicitly loads the Windows .NET Framework System.Windows.Forms assembly before this internal helper is called. No cross-platform claim is made.')]
    [CmdletBinding()]
    param($SyncObject, $Session, [int]$TimeoutSeconds)
    $monitor = $null
    $completed = $false
    $started = $false
    $attempt = 0
    $waiting = $false
    $state = $null
    $observations = [Collections.Generic.List[object]]::new()
    $timer = [Diagnostics.Stopwatch]::StartNew()
    try {
        while ($timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
            [Windows.Forms.Application]::DoEvents()
            if ($null -ne $monitor -and $monitor.ErrorCount -gt 0) {
                $script:SynchronizationError = $monitor.LastError
                Stop-Outcalget -Code 'OUTLOOK_SYNC_FAILED' -Message 'Outlook reported a synchronization error. No calendar export was published.' -ExitCode 3
            }
            Test-InternetConnection
            $state = Get-OutlookConnectionState -Session $Session
            Add-ConnectionObservation -State $state -Observations $observations -ElapsedMilliseconds $timer.ElapsedMilliseconds -Attempt $attempt
            if ($state.mode -in @(100, 200)) {
                Stop-Outcalget -Code 'OUTLOOK_CONNECTION_LOST' -Message ('Outlook is explicitly in Work Offline mode (primary Exchange mode {0}, Offline={1}). Connect Outlook normally and retry; the helper will not change that setting.' -f $state.mode, $state.offline) -ExitCode 3
            }
            if (-not $state.ready) {
                if ($null -ne $monitor) {
                    # Completion from a disconnected attempt is never reused after recovery.
                    $remainingMilliseconds = [Math]::Max(0, ([long]$TimeoutSeconds * 1000) - $timer.ElapsedMilliseconds)
                    Close-SynchronizationAttempt -SyncObject $SyncObject -Monitor $monitor -Started $started -TimeoutMilliseconds $remainingMilliseconds
                    $monitor = $null
                    $started = $false
                }
                if (-not $waiting) { Add-Trace -Phase 'sync-connection-wait' }
                $waiting = $true
            }
            elseif ($null -eq $monitor) {
                if ($attempt -ge 3) {
                    Stop-Outcalget -Code 'OUTLOOK_CONNECTION_UNSTABLE' -Message 'Outlook repeatedly lost its full-item Exchange connection. Three synchronization attempts were interrupted; no export was published.' -ExitCode 3
                }
                if ($waiting) { Add-Trace -Phase 'sync-connection-recovered' }
                $waiting = $false
                $monitor = [Outcalget.SyncMonitor]::new($SyncObject)
                $attempt++
                Add-ConnectionObservation -State $state -Observations $observations -ElapsedMilliseconds $timer.ElapsedMilliseconds -Attempt $attempt
                Add-Trace -Phase 'sync-start'
                $started = $true
                $SyncObject.Start()
                Add-Trace -Phase 'sync-wait'
            }
            elseif ($monitor.Completed) {
                Add-Trace -Phase 'sync-complete'
                $completed = $true
                return
            }
            Start-Sleep -Milliseconds 100
        }
        if ($null -ne $state -and -not $state.ready) {
            if ($state.mode -eq 500 -and -not $state.offline) {
                Stop-Outcalget -Code 'OUTLOOK_HEADERS_ONLY' -Message 'Outlook remained connected in headers-only mode (primary Exchange mode 500, Offline=False). Full-details export requires a full-item connection; no export was published.' -ExitCode 3
            }
            Stop-Outcalget -Code 'OUTLOOK_CONNECTION_LOST' -Message ('Outlook did not regain its full-item Exchange connection within the synchronization timeout (last primary Exchange mode {0}, Offline={1}). No export was published.' -f $state.mode, $state.offline) -ExitCode 3
        }
        Stop-Outcalget -Code 'OUTLOOK_SYNC_TIMEOUT' -Message 'No correlated synchronization completion was observed within the timeout. The cached calendar was not exported.' -ExitCode 7
    }
    finally {
        $script:SynchronizationConnectionObservations = @($observations.ToArray())
        if ($started -and -not $completed) {
            try { $SyncObject.Stop() }
            catch { $script:SyncCleanupFailed = $true }
        }
        if ($null -ne $monitor) {
            try { $monitor.Dispose() }
            catch { $script:SyncCleanupFailed = $true }
        }
    }
}

function ConvertTo-SafeConnectionObservation {
    param($Observations)
    if ($Observations -isnot [array] -or $Observations.Count -gt 24) {
        Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The worker returned invalid connection diagnostics.' -ExitCode 9
    }
    foreach ($item in $Observations) {
        foreach ($name in @('offline', 'primaryExchangeConnectionMode', 'elapsedMilliseconds', 'syncAttempt')) {
            if ($null -eq $item -or $item.PSObject.Properties.Name -notcontains $name) {
                Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The worker returned incomplete connection diagnostics.' -ExitCode 9
            }
        }
        if ($item.offline -isnot [bool] -or
            ($item.primaryExchangeConnectionMode -isnot [int] -and $item.primaryExchangeConnectionMode -isnot [long]) -or
            $item.primaryExchangeConnectionMode -lt [int]::MinValue -or $item.primaryExchangeConnectionMode -gt [int]::MaxValue -or
            ($item.elapsedMilliseconds -isnot [int] -and $item.elapsedMilliseconds -isnot [long]) -or
            $item.elapsedMilliseconds -lt 0 -or ($item.syncAttempt -isnot [int] -and $item.syncAttempt -isnot [long]) -or
            $item.syncAttempt -lt 0 -or $item.syncAttempt -gt 3) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The worker returned invalid connection diagnostic values.' -ExitCode 9
        }
        # Copy only numeric/boolean metadata, never extra fields from the reply.
        @{
            offline = $item.offline
            primaryExchangeConnectionMode = $item.primaryExchangeConnectionMode
            elapsedMilliseconds = $item.elapsedMilliseconds
            syncAttempt = $item.syncAttempt
        }
    }
}

function Invoke-CalendarSynchronization {
    # Local complexity exception: native setting checkpoint/restoration and RCW release share one failure-safe ownership transaction.
    param($Folder, $Session, [int]$TimeoutSeconds, [bool]$SharedDefaultCalendar = $false)
    $store = $null
    $objects = $null
    $sync = $null
    $previous = $null
    $changed = $false
    try {
        Add-Trace -Phase 'sync-prepare'
        $store = $Folder.Store
        # Microsoft documents Store=null for GetSharedDefaultFolder. The caller
        # permits this only after revalidating the pinned Exchange owner binding.
        if ($null -eq $store -and -not $SharedDefaultCalendar) {
            Stop-Outcalget -Code 'SOURCE_SYNC_UNSUPPORTED' -Message 'The selected calendar has no inspectable store and is not a validated shared-default Exchange binding. Select the shared default calendar by owner email when applicable.' -ExitCode 3
        }
        if ($null -ne $store -and [int]$store.ExchangeStoreType -eq 3) {
            Stop-Outcalget -Code 'SOURCE_SYNC_UNSUPPORTED' -Message 'The selected calendar is not in an Exchange store. Automatic server synchronization cannot be established for this source.' -ExitCode 3
        }
        Initialize-SynchronizationBridge
        $previous = [bool]$Folder.InAppFolderSyncObject
        $statePath = Join-Path -Path ([IO.Path]::GetDirectoryName($script:ProgressPath)) -ChildPath 'sync-state.json'
        $state = @{ runId = $script:RunId; originalInclusion = $previous; settingRestored = $previous }
        Write-AtomicJson -Path $statePath -Value $state
        if (-not $previous) {
            $changed = $true
            $Folder.InAppFolderSyncObject = $true
        }
        if (-not $Folder.InAppFolderSyncObject) {
            Stop-Outcalget -Code 'SYNC_FOLDER_NOT_INCLUDED' -Message 'The selected calendar could not be included in the native synchronization group.' -ExitCode 3
        }
        $objects = $Session.SyncObjects
        $sync = $objects.AppFolders
        Wait-CalendarSynchronization -SyncObject $sync -Session $Session -TimeoutSeconds $TimeoutSeconds
    }
    catch {
        # Capture the primary native stage before restoration records cleanup phases.
        if (-not $_.Exception.Data.Contains('OutcalgetPhase')) {
            $_.Exception.Data['OutcalgetPhase'] = $script:Phase
        }
        throw
    }
    finally {
        if ($changed) {
            try {
                $Folder.InAppFolderSyncObject = $previous
                if ([bool]$Folder.InAppFolderSyncObject -ne $previous) { $script:SyncCleanupFailed = $true }
                else {
                    $state.settingRestored = $true
                    Write-AtomicJson -Path $statePath -Value $state
                }
            }
            catch { $script:SyncCleanupFailed = $true }
            try { Add-Trace -Phase 'sync-restore' }
            catch { $script:SyncCleanupFailed = $true }
        }
        foreach ($reference in @($sync, $objects, $store)) {
            if ($null -ne $reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reference) }
                catch { $script:SyncCleanupFailed = $true }
            }
        }
    }
    if ($script:SyncCleanupFailed) {
        Stop-Outcalget -Code 'SYNC_CLEANUP_FAILED' -Message 'Synchronization finished, but its native folder setting or event subscription could not be restored safely.' -ExitCode 5
    }
}

function Restore-SynchronizationSetting {
    param($Folder, [string]$StatePath)
    Test-PrivateEntry -Path $StatePath
    $state = ConvertFrom-Json -InputObject (Read-BoundedText -Path $StatePath -MaxBytes 4096)
    if ($state.runId -cne $script:RunId -or $state.originalInclusion -isnot [bool] -or $state.settingRestored -isnot [bool]) {
        Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The synchronization recovery state is invalid.' -ExitCode 9
    }
    if (-not $state.settingRestored) {
        $Folder.InAppFolderSyncObject = $state.originalInclusion
        if ([bool]$Folder.InAppFolderSyncObject -ne $state.originalInclusion) {
            Stop-Outcalget -Code 'SYNC_CLEANUP_FAILED' -Message 'The original native synchronization folder setting could not be restored.' -ExitCode 5
        }
        $state.settingRestored = $true
        Write-AtomicJson -Path $StatePath -Value $state
        Add-Trace -Phase 'sync-restore'
    }
}

function Close-OutlookSession {
    # Local complexity exception: every native user-activity guard must complete before the single irreversible Quit call.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal cleanup of the verified process started by this export; never a pre-existing application.')]
    [CmdletBinding()]
    param($Application, $Lease)
    Test-OutlookLease -Lease $Lease
    $inspectors = $null
    $explorers = $null
    try {
        $inspectors = $Application.Inspectors
        $explorers = $Application.Explorers
        if ($inspectors.Count -gt 0 -or $explorers.Count -gt 1) {
            Stop-Outcalget -Code 'OUTLOOK_USER_ACTIVITY' -Message 'The started Outlook session has additional windows or open items. It was left open to protect user work.' -ExitCode 5
        }
        for ($index = 1; $index -le $explorers.Count; $index++) {
            $explorer = $null
            $inline = $null
            try {
                $explorer = $explorers.Item($index)
                $inline = $explorer.ActiveInlineResponse
                if ($null -ne $inline) {
                    Stop-Outcalget -Code 'OUTLOOK_USER_ACTIVITY' -Message 'The started Outlook session has an inline draft. It was left open to protect user work.' -ExitCode 5
                }
            }
            finally {
                foreach ($reference in @($inline, $explorer)) {
                    if ($null -ne $reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)) {
                        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reference) }
                        catch { $script:ComCleanupWarning = $true }
                    }
                }
            }
        }
        Test-OutlookLease -Lease $Lease
        Add-Trace -Phase 'outlook-close'
        $Application.Quit()
    }
    finally {
        foreach ($reference in @($explorers, $inspectors)) {
            if ($null -ne $reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reference) }
                catch { $script:ComCleanupWarning = $true }
            }
        }
    }
}

function Read-WorkerFeedback {
    # Local complexity exception: the correlated allowlist and optional native-start timestamp share one bounded progress read.
    param([string]$Directory, [string]$LastPhase)
    $path = Join-Path -Path $Directory -ChildPath 'progress.json'
    $phase = $null
    try {
        if ([IO.File]::Exists($path)) {
            $progress = ConvertFrom-Json -InputObject (Read-BoundedText -Path $path -MaxBytes 65536)
            $known = @('worker-request', 'outlook-attach', 'source-validate', 'calendar-picker',
                'outlook-ready-wait', 'outlook-setup-wait', 'outlook-online-wait', 'sync-prepare', 'sync-start', 'sync-wait',
                'sync-complete', 'sync-restore', 'exporter-create', 'exporter-settings',
                'ical-save', 'ical-save-complete', 'outlook-close', 'com-wait', 'com-acquired')
            if ($progress.runId -ceq $script:RunId -and $progress.phase -in $known) { $phase = $progress.phase }
            if ($progress.runId -ceq $script:RunId -and $progress.PSObject.Properties.Name -contains 'comAcquiredUtc' -and
                $null -ne $progress.comAcquiredUtc) {
                $captured = [datetime]::MinValue
                # Native JSON decoders can materialize ISO timestamps as DateTime.
                if ($progress.comAcquiredUtc -is [datetime]) {
                    $captured = $progress.comAcquiredUtc.ToUniversalTime()
                }
                else {
                    [void][datetime]::TryParseExact([string]$progress.comAcquiredUtc, 'o', [Globalization.CultureInfo]::InvariantCulture,
                        [Globalization.DateTimeStyles]::RoundtripKind, [ref]$captured)
                }
                if ($captured.Kind -eq [DateTimeKind]::Utc) {
                    $script:WorkerComAcquiredUtc = $captured
                }
            }
        }
    }
    catch { return $LastPhase }
    if ($null -ne $phase -and $phase -cne $LastPhase) {
        $script:LastWorkerPhase = $phase
        Write-RunRecord -Completed $false
        Write-PhaseMessage -Phase $phase
        return $phase
    }
    return $LastPhase
}

function Get-OutlookCoordinatorState {
    # Local complexity exception: exact marker shape and live PID/start-time/session checks fail closed as one ownership decision.
    param([string]$Path)
    if (-not [IO.File]::Exists($Path)) { return $null }
    Test-PrivateEntry -Path $Path
    try {
        $state = ConvertFrom-Json -InputObject (Read-BoundedText -Path $Path -MaxBytes 16384)
        $properties = @($state.PSObject.Properties.Name)
        $required = @('schemaVersion', 'runId', 'startedUtc', 'ownerId', 'ownerStartedTicks', 'phase', 'lease')
        $markerRun = [Guid]::Empty
        if ($properties.Count -ne 7 -or @($required | Where-Object { $_ -cnotin $properties }).Count -gt 0 -or
            $state.schemaVersion -ne 2 -or ($state.ownerId -isnot [int] -and $state.ownerId -isnot [long]) -or
            $state.ownerId -le 0 -or $state.ownerId -gt [int]::MaxValue -or
            ($state.ownerStartedTicks -isnot [long] -and $state.ownerStartedTicks -isnot [int]) -or $state.ownerStartedTicks -le 0 -or
            -not [Guid]::TryParseExact([string]$state.runId, 'D', [ref]$markerRun) -or $markerRun -eq [Guid]::Empty -or
            $state.phase -cnotin @('starting', 'ready', 'closing', 'uncertain')) {
            throw [InvalidOperationException]::new()
        }
        $owner = Get-ProcessIdentity -Id $state.ownerId
        $current = Get-ProcessIdentity -Id $PID
        if ($owner.startedTicks -ne $state.ownerStartedTicks -or $owner.sessionId -ne $current.sessionId) {
            throw [InvalidOperationException]::new()
        }
        return $state
    }
    catch {
        Stop-Outcalget -Code 'RECOVERY_REQUIRED' -Message 'An Outlook coordination marker is stale, invalid or owned by a terminated helper. Close Outlook normally and follow recovery instructions.' -ExitCode 5
    }
}

function Get-ProcessIdentity {
    param([int]$Id)
    $process = [Diagnostics.Process]::GetProcessById($Id)
    try {
        if ($process.HasExited) { throw [InvalidOperationException]::new() }
        return [pscustomobject]@{
            id = $process.Id
            startedTicks = $process.StartTime.ToUniversalTime().Ticks
            sessionId = $process.SessionId
        }
    }
    finally { $process.Dispose() }
}

function Set-OutlookCoordinatorState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private state checkpoint inside the approved lifecycle transaction; caller holds the lifecycle gate.')]
    [CmdletBinding()]
    param([string]$Phase, $Lease)
    $identity = Get-ProcessIdentity -Id $PID
    Write-AtomicJson -Path $script:LifecyclePending -Value @{
        schemaVersion = 2
        runId = $script:RunId
        startedUtc = $script:StartedUtc
        ownerId = $identity.id
        ownerStartedTicks = $identity.startedTicks
        phase = $Phase
        lease = $Lease
    }
}

function Open-OutlookUsageLease {
    param([bool]$RequireUnused)
    if ($null -ne $script:SessionUsageLease) { return }
    $key = Join-Path -Path $script:StateRoot -ChildPath 'outlook-users'
    $path = $null
    if ($RequireUnused) {
        $exclusive = Get-PathLock -Path $key -Kind 'session'
        try { $path = $exclusive.Name }
        finally { $exclusive.Dispose() }
    }
    else {
        # The lifecycle gate prevents another parent from racing initial file creation.
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $digest = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($key.ToUpperInvariant()))).Replace('-', '') }
        finally { $sha.Dispose() }
        $path = Join-Path -Path $script:LockRoot -ChildPath ('session-' + $digest + '.lock')
        if (-not [IO.File]::Exists($path)) {
            $created = Get-PathLock -Path $key -Kind 'session'
            $created.Dispose()
        }
        Test-PrivateEntry -Path $path
    }
    $script:SessionUsageLease = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
}

function Enter-OutlookCoordination {
    # Local complexity exception: the short gate makes process ownership, marker validation and shared-lease acquisition indivisible between cooperating parents.
    param([bool]$AllowStart, $Settings)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    Add-Trace -Phase 'coordination-wait'
    while ($timer.Elapsed.TotalSeconds -lt $CoordinationTimeoutSeconds) {
        $remaining = [Math]::Max(1, [int][Math]::Ceiling($CoordinationTimeoutSeconds - $timer.Elapsed.TotalSeconds))
        $gate = Wait-PathLock -Path (Join-Path -Path $script:StateRoot -ChildPath 'outlook-lifecycle') -Kind 'lifecycle' -TimeoutSeconds $remaining
        $script:LifecyclePending = Resolve-LocalPath -Path ($gate.Name + '.pending')
        $accepted = $false
        try {
            if ([IO.Directory]::Exists($script:LifecyclePending)) {
                Stop-Outcalget -Code 'RECOVERY_REQUIRED' -Message 'An Outlook recovery marker path is occupied unexpectedly.' -ExitCode 5
            }
            $state = Get-OutlookCoordinatorState -Path $script:LifecyclePending
            if ($null -ne $state -and $state.phase -ceq 'closing' -and $null -eq $script:SessionUsageLease) { continue }
            if ($null -ne $state -and $state.phase -cnotin @('ready', 'closing')) {
                Stop-Outcalget -Code 'RECOVERY_REQUIRED' -Message 'The earlier Outlook lifecycle is incomplete or uncertain. Follow recovery instructions before exporting.' -ExitCode 5
            }
            $instances = @(Get-OutlookInstance)
            if ($instances.Count -gt 1) {
                Stop-Outcalget -Code 'OUTLOOK_INSTANCE_AMBIGUOUS' -Message 'More than one Outlook Classic process exists in this session. No process will be started or closed.' -ExitCode 3
            }
            if ($null -eq $script:ActiveLifecycle) {
                $script:RunResult.outlookInitiallyRunning = $instances.Count -eq 1
            }
            $lifecycle = @{ lease = $null; startupTimeout = $StartupTimeoutSeconds; syncTimeout = $SyncTimeoutSeconds }
            if ($null -ne $state) {
                Test-OutlookLease -Lease $state.lease
                Open-OutlookUsageLease -RequireUnused $false
                $lifecycle.lease = $state.lease
                $script:RunResult.joinedOwnedSession = $state.runId -cne $script:RunId
                Add-Trace -Phase 'outlook-joined'
            }
            elseif ($instances.Count -eq 1) {
                Open-OutlookUsageLease -RequireUnused $false
                Add-Trace -Phase 'outlook-existing'
                $script:RunResult.synchronization = 'existing_session_not_requested'
            }
            elseif ($AllowStart) {
                try { Open-OutlookUsageLease -RequireUnused $true }
                catch {
                    if ($_.Exception.Data['OutcalgetCode'] -ceq 'RUN_BUSY') { continue }
                    throw
                }
                Set-OutlookCoordinatorState -Phase 'starting' -Lease $null
                $script:LifecyclePendingOwned = $true
                Add-Trace -Phase 'outlook-launch'
                $lifecycle.lease = Start-OutlookSession -ProfileName $Settings['profileName']
            }
            else {
                Stop-Outcalget -Code 'OUTLOOK_NOT_RUNNING' -Message 'Open Outlook Classic interactively before read-only checking.' -ExitCode 3
            }
            $script:ActiveLifecycle = $lifecycle
            $script:LifecycleLock = $gate
            $accepted = $true
            return $lifecycle
        }
        finally {
            if (-not $accepted) { $gate.Dispose(); Start-Sleep -Milliseconds 100 }
        }
    }
    Stop-Outcalget -Code 'COORDINATION_TIMEOUT' -Message 'Outlook is closing or another calendar export is still using its earlier session. Try again after it finishes.' -ExitCode 7
}

function Exit-OutlookCoordination {
    if ($null -ne $script:LifecycleLock) {
        $script:LifecycleLock.Dispose()
        $script:LifecycleLock = $null
    }
}

function Initialize-OutlookReadSession {
    $lifecycle = Enter-OutlookCoordination -AllowStart $false -Settings $null
    Exit-OutlookCoordination
    return $lifecycle
}

function Initialize-OutlookConfigurationSession {
    $lifecycle = Enter-OutlookCoordination -AllowStart $true -Settings @{ profileName = '' }
    try {
        $ready = Invoke-WorkerProcess -Mode 'setup-ready' -Settings $null -StagePath $null -TimeoutSeconds ($StartupTimeoutSeconds + 5) -Lifecycle $lifecycle
        if ([string]::IsNullOrWhiteSpace([string]$ready.profileName)) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The setup readiness worker did not identify an available Outlook profile.' -ExitCode 9
        }
        if ($script:LifecyclePendingOwned) {
            Set-OutlookCoordinatorState -Phase 'ready' -Lease $lifecycle.lease
        }
        return $lifecycle
    }
    finally { Exit-OutlookCoordination }
}

function Initialize-OutlookSession {
    param($Settings)
    if ($Settings['schemaVersion'] -eq 1 -or $Settings['sourceType'] -ceq 'folder') {
        $calendarKey = Get-CalendarFingerprint -StoreId $Settings['storeId'] -FolderId $Settings['folderEntryId']
        $script:CalendarLock = Get-PathLock -Path $calendarKey -Kind 'calendar'
    }
    else {
        $ownerKey = Get-CalendarFingerprint -StoreId $Settings['profileName'] -FolderId $Settings['sharedOwnerEntryId']
        $script:SharedSourceLock = Get-PathLock -Path $ownerKey -Kind 'shared-source'
    }
    $lifecycle = Enter-OutlookCoordination -AllowStart $true -Settings $Settings
    try {
        $source = Invoke-WorkerProcess -Mode 'ready' -Settings $Settings -StagePath $null -TimeoutSeconds ($StartupTimeoutSeconds + 5) -Lifecycle $lifecycle
        if ([string]::IsNullOrWhiteSpace([string]$source.storeId) -or [string]::IsNullOrWhiteSpace([string]$source.folderEntryId)) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The readiness worker did not return an identifiable calendar.' -ExitCode 9
        }
        $calendarKey = Get-CalendarFingerprint -StoreId ([string]$source.storeId) -FolderId ([string]$source.folderEntryId)
        if ($null -eq $script:CalendarLock) {
            $script:CalendarLock = Get-PathLock -Path $calendarKey -Kind 'calendar'
        }
        if ($null -ne $lifecycle.lease) {
            $syncData = Invoke-WorkerProcess -Mode 'sync' -Settings $Settings -StagePath $null -TimeoutSeconds ($SyncTimeoutSeconds + 15) -Lifecycle $lifecycle
            if ($syncData.synchronization -cne 'group_completed') {
                Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The synchronization worker did not confirm group completion.' -ExitCode 9
            }
            $script:RunResult.synchronization = 'group_completed'
        }
        if ($script:LifecyclePendingOwned) {
            Set-OutlookCoordinatorState -Phase 'ready' -Lease $lifecycle.lease
        }
        return $lifecycle
    }
    finally { Exit-OutlookCoordination }
}

function Restore-OutlookRunSetting {
    param($Settings)
    if ($null -eq $script:RunDirectory -or $null -eq $script:ActiveLifecycle -or $null -eq $script:ActiveLifecycle.lease) { return }
    $statePath = Join-Path -Path (Join-Path -Path $script:RunDirectory -ChildPath 'sync') -ChildPath 'sync-state.json'
    if (-not [IO.File]::Exists($statePath)) { return }
    $gate = Wait-PathLock -Path (Join-Path -Path $script:StateRoot -ChildPath 'outlook-lifecycle') -Kind 'lifecycle' -TimeoutSeconds $CoordinationTimeoutSeconds
    try {
        Test-PrivateEntry -Path $statePath
        $state = ConvertFrom-Json -InputObject (Read-BoundedText -Path $statePath -MaxBytes 4096)
        if ($state.runId -cne $script:RunId -or $state.settingRestored -isnot [bool]) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The synchronization recovery state is invalid.' -ExitCode 9
        }
        if (-not $state.settingRestored) {
            [void](Invoke-WorkerProcess -Mode 'restore' -Settings $Settings -StagePath $null -TimeoutSeconds 30 -Lifecycle $script:ActiveLifecycle)
        }
    }
    finally { $gate.Dispose() }
}

function Complete-OutlookSession {
    # Local complexity exception: owner quiescence, reference leases, normal Quit and uncertain-state preservation share a single cleanup lifetime.
    param($Settings)
    Exit-OutlookCoordination
    try { Restore-OutlookRunSetting -Settings $Settings }
    catch { $script:SyncCleanupFailed = $true; $script:KeepRunArtifacts = $true }
    finally {
        if ($null -ne $script:SessionUsageLease) {
            $script:SessionUsageLease.Dispose()
            $script:SessionUsageLease = $null
        }
    }
    if ($null -eq $script:LifecyclePending) { return }
    if (-not $script:OutlookStartAttempted -and -not $script:KeepRunArtifacts) {
        if ($script:LifecyclePendingOwned) {
            $gate = Wait-PathLock -Path (Join-Path -Path $script:StateRoot -ChildPath 'outlook-lifecycle') -Kind 'lifecycle' -TimeoutSeconds $CoordinationTimeoutSeconds
            try {
                $state = Get-OutlookCoordinatorState -Path $script:LifecyclePending
                if ($null -ne $state -and $state.runId -ceq $script:RunId -and $state.phase -ceq 'starting') {
                    [IO.File]::Delete($script:LifecyclePending)
                    $script:LifecyclePendingOwned = $false
                }
            }
            finally { $gate.Dispose() }
        }
        return
    }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    Add-Trace -Phase 'outlook-users-wait'
    while ($timer.Elapsed.TotalSeconds -lt $CoordinationTimeoutSeconds) {
        $remaining = [Math]::Max(1, [int][Math]::Ceiling($CoordinationTimeoutSeconds - $timer.Elapsed.TotalSeconds))
        $gate = Wait-PathLock -Path (Join-Path -Path $script:StateRoot -ChildPath 'outlook-lifecycle') -Kind 'lifecycle' -TimeoutSeconds $remaining
        $exclusive = $null
        try {
            if ($script:KeepRunArtifacts) {
                if ([IO.File]::Exists($script:LifecyclePending)) {
                    Test-PrivateEntry -Path $script:LifecyclePending
                    $state = ConvertFrom-Json -InputObject (Read-BoundedText -Path $script:LifecyclePending -MaxBytes 16384)
                    $state.phase = 'uncertain'
                    Write-AtomicJson -Path $script:LifecyclePending -Value $state
                }
                else { Set-OutlookCoordinatorState -Phase 'uncertain' -Lease $null }
                Stop-Outcalget -Code 'OUTLOOK_OUTCOME_UNCERTAIN' -Message 'Outlook was left open because native work or cleanup is uncertain. Follow recovery instructions before running another export.' -ExitCode 5
            }
            if ($null -eq $script:OwnedOutlookProcess -or $null -eq $script:OutlookLease) {
                Stop-Outcalget -Code 'OUTLOOK_OWNERSHIP_UNCERTAIN' -Message 'Startup ownership could not be established. Outlook was left untouched; inspect it and follow recovery instructions.' -ExitCode 5
            }
            $state = Get-OutlookCoordinatorState -Path $script:LifecyclePending
            if ($null -eq $state -or $state.runId -cne $script:RunId -or $state.phase -ceq 'uncertain') {
                Stop-Outcalget -Code 'OUTLOOK_OUTCOME_UNCERTAIN' -Message 'The shared Outlook session has uncertain ownership or work. It was left open for recovery.' -ExitCode 5
            }
            Set-OutlookCoordinatorState -Phase 'closing' -Lease $script:OutlookLease
            try {
                $exclusive = Get-PathLock -Path (Join-Path -Path $script:StateRoot -ChildPath 'outlook-users') -Kind 'session'
            }
            catch {
                if ($_.Exception.Data['OutcalgetCode'] -ceq 'RUN_BUSY') { continue }
                throw
            }
            if (-not $script:OwnedOutlookProcess.HasExited) {
                Test-OutlookLease -Lease $script:OutlookLease
                $lifecycle = @{ lease = $script:OutlookLease; startupTimeout = 15; syncTimeout = 30 }
                [void](Invoke-WorkerProcess -Mode 'close' -Settings $Settings -StagePath $null -TimeoutSeconds 30 -Lifecycle $lifecycle)
                if (-not $script:OwnedOutlookProcess.WaitForExit(30000)) {
                    Stop-Outcalget -Code 'OUTLOOK_CLOSE_TIMEOUT' -Message 'The owned Outlook session did not exit after normal Quit. No forced close was attempted.' -ExitCode 5
                }
            }
            $script:RunResult.outlookClosed = $true
            Add-Trace -Phase 'outlook-closed'
            [IO.File]::Delete($script:LifecyclePending)
            $script:LifecyclePendingOwned = $false
            return
        }
        finally {
            if ($null -ne $exclusive) { $exclusive.Dispose() }
            $gate.Dispose()
            Start-Sleep -Milliseconds 100
        }
    }
    Stop-Outcalget -Code 'OUTLOOK_USERS_TIMEOUT' -Message 'Other export runs are still using the Outlook session started here. It was left open; follow recovery instructions after all runs finish.' -ExitCode 5
}

function Invoke-OutlookWorker {
    # Local size/complexity exception: preserve the working COM acquisition, export, reply and release transaction.
    param([string]$Directory, [string]$Token)
    $application = $null
    $session = $null
    $folder = $null
    $exporter = $null
    $replyPath = $null
    $reply = $null
    $comLock = $null
    $comPending = $null
    $comPendingOwned = $false
    try {
        Test-HostEnvironment
        Test-DesktopSession
        $directoryFull = Resolve-LocalPath -Path $Directory
        $expectedRuns = Join-Path -Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath 'osp\consulthelp\outcalget\runs'
        if (-not $directoryFull.StartsWith($expectedRuns + '\', [StringComparison]::OrdinalIgnoreCase)) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'Invalid internal worker directory.' -ExitCode 9
        }
        Test-PrivateEntry -Path $directoryFull
        $parentDirectory = [IO.Path]::GetDirectoryName($directoryFull)
        $directoryRunId = [IO.Path]::GetFileName($parentDirectory)
        $directoryMode = [IO.Path]::GetFileName($directoryFull)
        $parsedRun = [Guid]::Empty
        if ([IO.Path]::GetDirectoryName($parentDirectory) -ine $expectedRuns -or
            -not [Guid]::TryParseExact($directoryRunId, 'D', [ref]$parsedRun) -or $directoryMode -notin @('select', 'select-shared', 'nav-list', 'select-nav', 'check', 'setup-ready', 'ready', 'sync', 'restore', 'export', 'close')) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'Invalid internal worker directory layout.' -ExitCode 9
        }
        $script:RunId = $directoryRunId
        $replyPath = Join-Path -Path $directoryFull -ChildPath 'result.json'
        $script:ProgressPath = Join-Path -Path $directoryFull -ChildPath 'progress.json'
        $reply = @{
            runId       = $script:RunId
            mode        = $directoryMode
            ok          = $false
            error       = $null
            data        = $null
            diagnostics = @{
                stageOwned = $false
            }
        }
        Add-Trace -Phase 'worker-request'
        $requestPath = Join-Path -Path $directoryFull -ChildPath 'request.json'
        Test-PrivateEntry -Path $requestPath
        $request = ConvertFrom-Json -InputObject (Read-BoundedText -Path $requestPath -MaxBytes 131072)
        if ($request.token -cne $Token -or $request.runId -cne $directoryRunId -or
            $request.mode -cne $directoryMode) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'Internal request validation failed.' -ExitCode 9
        }
        if (($request.queueTimeoutSeconds -isnot [int] -and $request.queueTimeoutSeconds -isnot [long]) -or $request.queueTimeoutSeconds -lt 15 -or $request.queueTimeoutSeconds -gt 1800) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The internal coordination timeout is invalid.' -ExitCode 9
        }
        $script:StateRoot = [IO.Path]::GetDirectoryName($expectedRuns)
        $script:LockRoot = Join-Path -Path $script:StateRoot -ChildPath 'locks'
        Add-Trace -Phase 'com-wait'
        $comLock = Wait-PathLock -Path (Join-Path -Path $script:StateRoot -ChildPath 'outlook-com') -Kind 'com' -TimeoutSeconds ([int]$request.queueTimeoutSeconds)
        $script:ComAcquiredUtc = [datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        Add-Trace -Phase 'com-acquired'
        $comPending = Resolve-LocalPath -Path ($comLock.Name + '.pending')
        if ([IO.File]::Exists($comPending) -or [IO.Directory]::Exists($comPending)) {
            Stop-Outcalget -Code 'RECOVERY_REQUIRED' -Message 'A prior Outlook COM worker has an uncertain outcome. Close Outlook normally and follow pending-marker recovery instructions.' -ExitCode 5
        }
        Write-AtomicJson -Path $comPending -Value @{ runId = $script:RunId; mode = $request.mode; startedUtc = $script:StartedUtc }
        $comPendingOwned = $true
        # A terminated COM worker can leave Outlook still working. Never start another
        # native operation behind an uncertain session marker, even for another output.
        $sha = [Security.Cryptography.SHA256]::Create()
        $lifecycleKey = (Join-Path -Path $script:StateRoot -ChildPath 'outlook-lifecycle').ToUpperInvariant()
        try { $digest = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($lifecycleKey))).Replace('-', '') }
        finally { $sha.Dispose() }
        $marker = Join-Path -Path $script:LockRoot -ChildPath ('lifecycle-' + $digest + '.lock.pending')
        if ([IO.File]::Exists($marker)) {
            $coordination = Get-OutlookCoordinatorState -Path $marker
            if ($coordination.phase -ceq 'uncertain') {
                Stop-Outcalget -Code 'RECOVERY_REQUIRED' -Message 'Earlier Outlook work has an uncertain outcome. Follow recovery instructions before another native operation.' -ExitCode 5
            }
        }
        Add-Trace -Phase 'outlook-attach'
        $setupTimer = [Diagnostics.Stopwatch]::StartNew()
        if ($request.mode -in @('setup-ready', 'ready', 'sync', 'restore', 'close')) {
            if ($null -eq $request.lifecycle -or $request.lifecycle.startupTimeout -notin (15..600) -or
                $request.lifecycle.syncTimeout -notin (30..600)) {
                Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The internal Outlook lifecycle request is invalid.' -ExitCode 9
            }
            if ($request.mode -notin @('setup-ready', 'ready') -or $null -ne $request.lifecycle.lease) {
                Test-OutlookLease -Lease $request.lifecycle.lease
            }
            $application = Get-ReadyOutlookApplication -TimeoutSeconds ([int]$request.lifecycle.startupTimeout)
        }
        else {
            $application = Invoke-ComRead -Action {
                [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
            }
        }
        $reply.diagnostics.outlookVersion = [string]$application.Version
        if ($request.mode -eq 'close') {
            Close-OutlookSession -Application $application -Lease $request.lifecycle.lease
            $reply.data = @{ warnings = @() }
        }
        else {
            $session = Invoke-ComRead -Action { $application.GetNamespace('MAPI') }
            Add-Trace -Phase 'source-validate'
        }
        if ($request.mode -eq 'setup-ready') {
            $remaining = [Math]::Max(0, [int][Math]::Floor($request.lifecycle.startupTimeout - $setupTimer.Elapsed.TotalSeconds))
            $profileName = Get-ReadyOutlookSetup -Session $session -RequireOnline ($null -ne $request.lifecycle.lease) -TimeoutSeconds $remaining
            if ($null -ne $request.lifecycle.lease) { Test-OutlookLease -Lease $request.lifecycle.lease }
            $reply.data = @{ profileName = $profileName; warnings = @() }
        }
        elseif ($request.mode -eq 'nav-list') {
            $reply.data = @{
                calendars = @(Get-NavigationCalendar -Application $application -Fingerprint '')
                unavailableCount = $script:NavigationUnavailableCount
            }
        }
        elseif ($request.mode -in @('select', 'select-shared', 'select-nav')) {
            Add-Trace -Phase 'calendar-picker'
            $sharedBinding = $null
            if ($request.mode -eq 'select') { $folder = $session.PickFolder() }
            elseif ($request.mode -eq 'select-shared') {
                $sharedBinding = Get-SharedCalendarBinding -Session $session -Owner ([string]$request.configuration.owner) -ExpectedOwnerId ''
                $folder = $sharedBinding.Folder
            }
            else {
                $fingerprint = [string]$request.configuration.fingerprint
                if ($fingerprint -cnotmatch '^[A-F0-9]{64}$') {
                    Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The navigation selection token is invalid.' -ExitCode 9
                }
                $folder = Get-NavigationCalendar -Application $application -Fingerprint $fingerprint
            }
            if ($null -eq $folder) {
                Stop-Outcalget -Code 'CANCELLED' -Message 'Calendar selection was cancelled.' -ExitCode 8
            }
            if ([int]$folder.DefaultItemType -ne 1) {
                Stop-Outcalget -Code 'NOT_A_CALENDAR' -Message 'Choose a calendar folder in Outlook Classic.' -ExitCode 2
            }
            $reply.data = @{
                sourceType = 'folder'
                profileName = [string]$session.CurrentProfileName
                storeId = [string]$folder.StoreID
                folderEntryId = [string]$folder.EntryID
                displayPath = [string]$folder.FolderPath
            }
            if ($null -ne $sharedBinding) {
                $reply.data.sourceType = 'sharedDefaultCalendar'
                $reply.data.sharedCalendarOwner = $sharedBinding.Owner
                $reply.data.sharedOwnerEntryId = $sharedBinding.OwnerEntryId
            }
        }
        elseif ($request.mode -notin @('close', 'nav-list')) {
            $settings = Read-FlatJson -Text (ConvertTo-Json -InputObject $request.configuration -Compress)
            Test-Configuration -Values $settings
            if ($settings['userSid'] -cne $script:UserSid) {
                Stop-Outcalget -Code 'PROFILE_MISMATCH' -Message 'The Windows user or running Outlook profile does not match the saved configuration.' -ExitCode 2
            }
            if ($request.mode -eq 'ready' -and $null -ne $request.lifecycle.lease) {
                if ($settings['schemaVersion'] -eq 2 -and $settings['sourceType'] -ceq 'sharedDefaultCalendar') {
                    Wait-OutlookConnection -Session $session -TimeoutSeconds ([int]$request.lifecycle.startupTimeout)
                }
                $folder = Get-ReadyOutlookSource -Session $session -Settings $settings -TimeoutSeconds ([int]$request.lifecycle.startupTimeout)
            }
            else {
                if ($settings['profileName'] -cne [string]$session.CurrentProfileName) {
                    Stop-Outcalget -Code 'PROFILE_MISMATCH' -Message 'The Windows user or running Outlook profile does not match the saved configuration.' -ExitCode 2
                }
                $folder = Invoke-ComRead -Action {
                    Resolve-ConfiguredCalendar -Session $session -Settings $settings
                }
            }
            if ([int]$folder.DefaultItemType -ne 1) {
                Stop-Outcalget -Code 'SOURCE_MISMATCH' -Message 'The resolved calendar does not match the configured binding.' -ExitCode 2
            }
            $reply.diagnostics.bindingMatched = $true
            $connection = $null
            try {
                $connection = [int]$session.ExchangeConnectionMode
            }
            catch {
                $reply.diagnostics.connectionModeUnavailable = $true
            }
            $reply.diagnostics.primaryExchangeConnectionMode = $connection
            $reply.data = @{
                warnings = @()
            }
            if ($connection -in @(100, 200, 300, 400)) {
                $reply.data.warnings = @(@{
                        code    = 'PRIMARY_EXCHANGE_OFFLINE'
                        message = 'The primary Exchange account reports offline/disconnected status; selected-calendar freshness is unverified.'
                    })
            }
            if ($request.mode -eq 'ready') {
                $reply.data.storeId = [string]$folder.StoreID
                $reply.data.folderEntryId = [string]$folder.EntryID
                Test-InternetConnection
                if ($null -ne $request.lifecycle.lease) {
                    Wait-OutlookConnection -Session $session -TimeoutSeconds ([int]$request.lifecycle.startupTimeout)
                    Test-OutlookLease -Lease $request.lifecycle.lease
                }
            }
            elseif ($request.mode -eq 'sync') {
                Test-InternetConnection
                $sharedDefaultCalendar = $settings['schemaVersion'] -eq 2 -and $settings['sourceType'] -ceq 'sharedDefaultCalendar'
                Invoke-CalendarSynchronization -Folder $folder -Session $session -TimeoutSeconds ([int]$request.lifecycle.syncTimeout) -SharedDefaultCalendar $sharedDefaultCalendar
                Test-OutlookLease -Lease $request.lifecycle.lease
                $reply.data.synchronization = 'group_completed'
            }
            elseif ($request.mode -eq 'restore') {
                $statePath = Join-Path -Path (Join-Path -Path $parentDirectory -ChildPath 'sync') -ChildPath 'sync-state.json'
                Restore-SynchronizationSetting -Folder $folder -StatePath $statePath
                Test-OutlookLease -Lease $request.lifecycle.lease
            }
            if ($request.mode -eq 'export') {
                $output = Resolve-LocalPath -Path $settings['outputPath']
                $expectedStage = Join-Path -Path ([IO.Path]::GetDirectoryName($output)) -ChildPath ('.outcalget-' + $script:RunId + '.partial.ics')
                if ($request.stagePath -cne $expectedStage -or [IO.File]::Exists($expectedStage) -or [IO.Directory]::Exists($expectedStage)) {
                    Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The internal staging path is invalid or already occupied.' -ExitCode 9
                }
                Test-PrivateEntry -Path ([IO.Path]::GetDirectoryName($output))
                $start = [datetime]::ParseExact($request.startDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
                if ($request.endDate -cne $start.AddDays(29).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)) {
                    Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The internal date range is invalid.' -ExitCode 9
                }
                Add-Trace -Phase 'exporter-create'
                $exporter = Invoke-ComRead -Action {
                    $folder.GetCalendarExporter()
                }
                Add-Trace -Phase 'exporter-source-validate'
                Test-CalendarExporterSource -Exporter $exporter -Folder $folder -Settings $settings -Session $session
                Add-Trace -Phase 'exporter-settings'
                $exporter.CalendarDetail = 2
                $exporter.IncludeWholeCalendar = $false
                $exporter.StartDate = $start
                $exporter.EndDate = $start.AddDays(29)
                $exporter.IncludePrivateDetails = [bool]$settings['includePrivateDetails']
                $exporter.IncludeAttachments = [bool]$settings['includeAttachments']
                $exporter.RestrictToWorkingHours = $false
                Test-InternetConnection
                if ($null -ne $request.lifecycle -and $null -ne $request.lifecycle.lease) {
                    Test-OutlookLease -Lease $request.lifecycle.lease
                    if ($session.Offline -or [int]$session.ExchangeConnectionMode -notin @(600, 700, 800)) {
                        Stop-Outcalget -Code 'OUTLOOK_CONNECTION_LOST' -Message 'Outlook lost its full-item Exchange connection before export.' -ExitCode 3
                    }
                }
                Add-Trace -Phase 'ical-save'
                # The stage was vacant when this worker accepted it. Only a worker that
                # reached the native write may authorise parent cleanup of this pathname.
                $reply.diagnostics.stageOwned = $true
                $exporter.SaveAsICal($expectedStage)
                Add-Trace -Phase 'ical-save-complete'
            }
        }
        $reply.ok = $true
    }
    catch {
        if ($null -ne $reply) {
            $reply.error = Get-ExceptionInfo -Record $_
        }
    }
    finally {
        foreach ($reference in @($exporter, $folder, $session, $application)) {
            if ($null -ne $reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)) {
                try {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reference)
                }
                catch {
                    $reply.diagnostics.comReleaseFailed = $true
                }
            }
        }
        try {
            if ($null -ne $reply -and $null -ne $replyPath) {
                $reply.diagnostics.syncCleanupFailed = $script:SyncCleanupFailed
                $reply.diagnostics.syncErrorCode = $script:SynchronizationError
                if ($script:SynchronizationConnectionObservations.Count -gt 0) {
                    $reply.diagnostics.connectionObservations = @($script:SynchronizationConnectionObservations)
                }
                $reply.diagnostics.comReleaseFailed = $script:ComCleanupWarning -or
                    ($reply.diagnostics.ContainsKey('comReleaseFailed') -and $reply.diagnostics.comReleaseFailed)
                $reply.diagnostics.events = @($script:Trace.ToArray())
                Write-AtomicJson -Path $replyPath -Value $reply
                if ($comPendingOwned) { [IO.File]::Delete($comPending) }
            }
        }
        finally { if ($null -ne $comLock) { $comLock.Dispose() } }
    }
    if ($null -eq $reply -or -not $reply.ok) {
        return 1
    }
    return 0
}

function Invoke-WorkerProcess {
    # Local size/complexity exception: process, timeout, reply validation and disposal share one ownership lifetime.
    param([string]$Mode, $Settings, [string]$StagePath, [int]$TimeoutSeconds, $Lifecycle = $null)
    $directory = New-PrivateDirectory -Path (Join-Path -Path $script:RunDirectory -ChildPath $Mode)
    $token = [Guid]::NewGuid().ToString('D')
    $request = @{
        token         = $token
        runId         = $script:RunId
        mode          = $Mode
        configuration = $Settings
        stagePath     = $StagePath
        startDate     = $script:Window.start.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        endDate       = $script:Window.end.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        lifecycle     = $Lifecycle
        queueTimeoutSeconds = $CoordinationTimeoutSeconds
    }
    Write-AtomicJson -Path (Join-Path -Path $directory -ChildPath 'request.json') -Value $request
    $arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-File', $script:ScriptPath,
        '-Worker', '-WorkDirectory', $directory, '-WorkerToken', $token)
    $info = [Diagnostics.ProcessStartInfo]::new()
    $info.FileName = Join-Path -Path $PSHOME -ChildPath 'powershell.exe'
    $info.Arguments = (($arguments | ForEach-Object {
                ConvertTo-NativeArgument -Value $_
            }) -join ' ')
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $info
    $stdoutDrain = $null
    $stderrDrain = $null
    $started = $false
    $finished = $false
    $receivedResponse = $false
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $launchUtc = [datetime]::UtcNow
    $script:WorkerComAcquiredUtc = $null
    $lastProgress = $null
    $progressPoll = [Diagnostics.Stopwatch]::StartNew()
    try {
        Add-Trace -Phase 'worker-start'
        $started = $process.Start()
        if (-not $started) {
            Stop-Outcalget -Code 'WORKER_START_FAILED' -Message 'The Outlook worker could not be started.' -ExitCode 3
        }
        # Drain to a null stream: never retain or forward raw host/COM error text.
        $stdoutDrain = $process.StandardOutput.BaseStream.CopyToAsync([IO.Stream]::Null)
        $stderrDrain = $process.StandardError.BaseStream.CopyToAsync([IO.Stream]::Null)
        while (-not $process.WaitForExit(100)) {
            if ($progressPoll.ElapsedMilliseconds -ge 500) {
                $lastProgress = Read-WorkerFeedback -Directory $directory -LastPhase $lastProgress
                $progressPoll.Restart()
            }
            $nativeSeconds = 0
            if ($null -ne $script:WorkerComAcquiredUtc) {
                $nativeSeconds = $timer.Elapsed.TotalSeconds - ($script:WorkerComAcquiredUtc - $launchUtc).TotalSeconds
            }
            if ($nativeSeconds -ge $TimeoutSeconds -or $timer.Elapsed.TotalSeconds -ge ($TimeoutSeconds + $CoordinationTimeoutSeconds + 10)) {
                $script:KeepRunArtifacts = $true
                Stop-Outcalget -Code 'OUTLOOK_TIMEOUT' -Message 'Outlook did not finish within the worker timeout. Check Outlook for a dialogue. Any previous export was not refreshed.' -ExitCode 7
            }
        }
        $finished = $true
        [void](Read-WorkerFeedback -Directory $directory -LastPhase $lastProgress)
        $script:WorkerDiagnostics.Add(@{
                mode                = $Mode
                processExitCode     = $process.ExitCode
                elapsedMilliseconds = $timer.ElapsedMilliseconds
            })
        $resultPath = Join-Path -Path $directory -ChildPath 'result.json'
        if (-not [IO.File]::Exists($resultPath)) {
            Stop-Outcalget -Code 'WORKER_NO_RESULT' -Message 'The worker stopped without a result. Policy, startup or a host error may have blocked it; inspect the debug report.' -ExitCode 3
        }
        $replyLimit = 65536
        if ($Mode -eq 'nav-list') { $replyLimit = 131072 }
        $reply = ConvertFrom-Json -InputObject (Read-BoundedText -Path $resultPath -MaxBytes $replyLimit)
        if ($reply.runId -cne $script:RunId -or $reply.mode -cne $Mode -or $reply.ok -isnot [bool]) {
            Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The worker response could not be correlated to this run.' -ExitCode 9
        }
        if ($Mode -eq 'export') {
            $ownershipPresent = $reply.diagnostics.PSObject.Properties.Name -contains 'stageOwned'
            if ($ownershipPresent -and $reply.diagnostics.stageOwned -isnot [bool]) {
                Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The worker response has invalid staging ownership.' -ExitCode 9
            }
            $script:StageOwned = $ownershipPresent -and $reply.diagnostics.stageOwned
            if ($reply.ok -and -not $script:StageOwned) {
                Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The worker did not confirm staging ownership.' -ExitCode 9
            }
        }
        $receivedResponse = $true
        # Only explicitly safe diagnostics are copied; source selection data stays private.
        $safe = @{
            mode   = $Mode
            events = @($reply.diagnostics.events)
        }
        if ($reply.diagnostics.PSObject.Properties.Name -contains 'outlookVersion') {
            $safe.outlookVersion = $reply.diagnostics.outlookVersion
        }
        if ($reply.diagnostics.PSObject.Properties.Name -contains 'primaryExchangeConnectionMode') {
            $safe.primaryExchangeConnectionMode = $reply.diagnostics.primaryExchangeConnectionMode
        }
        if ($reply.diagnostics.PSObject.Properties.Name -contains 'connectionObservations') {
            $safe.connectionObservations = @(ConvertTo-SafeConnectionObservation -Observations $reply.diagnostics.connectionObservations)
        }
        foreach ($name in @('syncCleanupFailed', 'syncErrorCode', 'comReleaseFailed')) {
            if ($reply.diagnostics.PSObject.Properties.Name -contains $name) {
                $value = $reply.diagnostics.$name
                if (($name -eq 'syncErrorCode' -and $null -ne $value -and $value -isnot [int]) -or
                    ($name -ne 'syncErrorCode' -and $value -isnot [bool])) {
                    Stop-Outcalget -Code 'WORKER_PROTOCOL' -Message 'The worker returned an invalid lifecycle diagnostic.' -ExitCode 9
                }
                $safe[$name] = $value
            }
        }
        if ($safe.ContainsKey('syncCleanupFailed') -and $safe.syncCleanupFailed) {
            $script:SyncCleanupFailed = $true
            $script:KeepRunArtifacts = $true
        }
        if ($safe.ContainsKey('comReleaseFailed') -and $safe.comReleaseFailed) { $script:ComCleanupWarning = $true }
        $script:WorkerDiagnostics.Add($safe)
        if (-not $reply.ok) {
            $script:RemoteError = $reply.error
            Stop-Outcalget -Code $reply.error.code -Message $reply.error.message -ExitCode ([int]$reply.error.exitCode)
        }
        if ($process.ExitCode -ne 0) {
            Stop-Outcalget -Code 'WORKER_EXIT_FAILED' -Message 'The worker returned an unsuccessful process exit code.' -ExitCode 4
        }
        return $reply.data
    }
    finally {
        if ($Mode -eq 'export' -and $started -and -not $receivedResponse) {
            $script:KeepRunArtifacts = $true
        }
        if ($started -and -not $finished) {
            $script:KeepRunArtifacts = $true
            try {
                if (-not $process.HasExited) {
                    $process.Kill()
                    [void]$process.WaitForExit(2000)
                }
            }
            catch {
                $script:WorkerDiagnostics.Add(@{
                        mode             = $Mode
                        workerStopFailed = $true
                    })
            }
        }
        $progressPath = Join-Path -Path $directory -ChildPath 'progress.json'
        if ([IO.File]::Exists($progressPath)) {
            try {
                $progress = ConvertFrom-Json -InputObject (Read-BoundedText -Path $progressPath -MaxBytes 65536)
                if ($progress.runId -ceq $script:RunId) {
                    $script:WorkerDiagnostics.Add(@{
                            mode      = $Mode
                            lastPhase = $progress.phase
                            events    = @($progress.events)
                        })
                }
            }
            catch {
                $script:WorkerDiagnostics.Add(@{
                        mode                = $Mode
                        progressUnavailable = $true
                    })
            }
        }
        if ($started) {
            try {
                $drained = [Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($stdoutDrain, $stderrDrain), 1000)
                if (-not $drained) {
                    $script:WorkerDiagnostics.Add(@{
                            mode            = $Mode
                            drainIncomplete = $true
                        })
                }
            }
            catch {
                $script:WorkerDiagnostics.Add(@{
                        mode             = $Mode
                        drainUnavailable = $true
                    })
            }
        }
        $timer.Stop()
        $process.Dispose()
    }
}

function Get-EnvironmentDiagnostic {
    $data = @{
        osVersion         = [Environment]::OSVersion.Version.ToString()
        is64BitOS         = [Environment]::Is64BitOperatingSystem
        is64BitProcess    = [Environment]::Is64BitProcess
        powershellVersion = $PSVersionTable.PSVersion.ToString()
        powershellEdition = $PSVersionTable.PSEdition
        clrVersion        = [Environment]::Version.ToString()
        languageMode      = [string]$ExecutionContext.SessionState.LanguageMode
        apartment         = [string][Threading.Thread]::CurrentThread.GetApartmentState()
    }
    try {
        $data.executionPolicies = @(Get-ExecutionPolicy -List | ForEach-Object {
                @{
                    scope  = [string]$_.Scope
                    policy = [string]$_.ExecutionPolicy
                }
            })
    }
    catch {
        $data.executionPolicies = @()
    }
    try {
        $data.signatureStatus = [string](Get-AuthenticodeSignature -LiteralPath $script:ScriptPath).Status
    }
    catch {
        $data.signatureStatus = 'unavailable'
    }
    try {
        $data.scriptSha256 = (Get-FileHash -LiteralPath $script:ScriptPath -Algorithm SHA256).Hash
    }
    catch {
        $data.scriptSha256 = $null
    }
    try {
        $sessionId = [Diagnostics.Process]::GetCurrentProcess().SessionId
        $data.outlookProcessCountInSession = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Where-Object {
                $_.SessionId -eq $sessionId
            }).Count
    }
    catch {
        $data.outlookProcessCountInSession = $null
    }
    return $data
}

function New-DiagnosticReport {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Pure in-memory record construction; no side effects.')]
    [CmdletBinding()]
    param($Result)
    $report = @{
        diagnosticSchemaVersion    = 1
        tool                       = 'outcalget'
        version                    = '0.3.5'
        runId                      = $script:RunId
        timestampUtc               = [datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        environment                = $script:EnvironmentDiagnostics
        operation                  = $Result.operation
        status                     = $Result.status
        exitCode                   = $Result.exitCode
        published                  = $Result.published
        elapsedMilliseconds        = $Result.elapsedMilliseconds
        lastParentPhase            = $script:Phase
        events                     = @($script:Trace.ToArray())
        workers                    = @($script:WorkerDiagnostics.ToArray())
        settingsSummary            = $script:SettingsSummary
        lastWorkerPhase            = $script:LastWorkerPhase
        lifecycle                  = @{
            internetAvailable       = $Result.internetAvailable
            outlookInitiallyRunning = $Result.outlookInitiallyRunning
            outlookStarted          = $Result.outlookStarted
            outlookClosed           = $Result.outlookClosed
            synchronization         = $Result.synchronization
            freshness               = $Result.freshness
        }
        error                      = $Result.error
        failureDetails             = $script:FailureDetails
        temporaryArtifactsRetained = $script:KeepRunArtifacts
        privacy                    = 'No calendar content, mailbox/folder IDs, usernames, profile names, paths, credentials or raw exception messages are collected.'
    }
    return $report
}

function Get-SynchronizationSummary {
    param([string]$Synchronization)
    switch ($Synchronization) {
        'group_completed' {
            return 'Outlook synchronization completed. Server/cache equivalence was not independently checked.'
        }
        'existing_session_not_requested' {
            return 'Used the existing Outlook session; explicit synchronization was not requested.'
        }
        default { return 'Outlook synchronization was not requested or did not complete.' }
    }
}

function New-RunRecord {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Pure in-memory record construction; no side effects.')]
    [CmdletBinding()]
    param([bool]$Completed)
    $result = $script:RunResult
    $failurePhase = $null
    if ($null -ne $script:FailureDetails -and $script:FailureDetails.ContainsKey('phase')) {
        $failurePhase = $script:FailureDetails.phase
    }
    $record = [ordered]@{
        logSchemaVersion           = 1
        tool                       = 'outcalget'
        version                    = '0.3.5'
        runId                      = $script:RunId
        startedUtc                 = $script:StartedUtc
        updatedUtc                 = [datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        completed                  = $Completed
        scheduled                  = $script:ScheduledRun
        operation                  = $result.operation
        status                     = 'running'
        exitCode                   = $null
        published                  = $result.published
        elapsedMilliseconds        = $script:Clock.ElapsedMilliseconds
        lastParentPhase            = $script:Phase
        lastWorkerPhase            = $script:LastWorkerPhase
        failurePhase               = $failurePhase
        internetAvailable          = $result.internetAvailable
        outlookInitiallyRunning    = $result.outlookInitiallyRunning
        outlookStarted             = $result.outlookStarted
        outlookClosed              = $result.outlookClosed
        namedProfile               = $script:NamedExportProfile
        sourceType                 = $result.sourceType
        joinedOwnedSession         = $result.joinedOwnedSession
        synchronization            = $result.synchronization
        synchronizationSummary     = Get-SynchronizationSummary -Synchronization $result.synchronization
        freshness                  = $result.freshness
        errorCode                  = $null
        warningCodes               = @($result.warnings | ForEach-Object {
                $_.code
            })
        temporaryArtifactsRetained = $script:KeepRunArtifacts
        debug                      = $null
    }
    if ($Completed) {
        $record.status = $result.status
        $record.exitCode = $result.exitCode
        if ($null -ne $result.error) {
            $record.errorCode = $result.error.code
        }
    }
    if ($script:DiagnosticsRequested) {
        try {
            $record.debug = New-DiagnosticReport -Result $result
        }
        catch {
            # Optional enrichment cannot prevent required minimal operational logging.
            $record.debug = @{
                unavailable = $true
                code        = 'DEBUG_METADATA_UNAVAILABLE'
            }
        }
    }
    return $record
}

function Write-RunRecord {
    param([bool]$Completed)
    $record = New-RunRecord -Completed $Completed
    $text = ConvertTo-Json -InputObject $record -Depth 14 -Compress
    if ([Text.Encoding]::UTF8.GetByteCount($text) -gt 65536) {
        # Keep a bounded diagnostic trace; never lose the minimal outcome to optional detail.
        $record.debug = @{
            unavailable = $true
            code        = 'DEBUG_SIZE_LIMIT'
        }
    }
    try {
        Write-AtomicJson -Path $script:RunLogPath -Value $record
    }
    catch {
        Stop-Outcalget -Code 'RUN_LOG_WRITE_FAILED' -Message 'Required operational logging failed. Inspect the last run record; no later business step was attempted.' -ExitCode 5
    }
}

function Remove-ExpiredRunRecord {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal operational-log retention; disabled for preview and guarded by active handles.')]
    [CmdletBinding()]
    param([string]$Directory)
    # The per-run .active handle prevents deletion of any still-running parent's record.
    # Only this exact tool-owned name format is eligible. Sensitive runs/staging are NEVER pruned here.
    $cutoff = [datetime]::UtcNow.AddDays(-14)
    $files = @(Get-ChildItem -LiteralPath $Directory -Filter 'outcalget-*.json' -File |
            Where-Object {
                $_.Name -match '^outcalget-\d{8}T\d{6}Z-[0-9a-f-]{36}\.json$'
            } |
            Sort-Object -Property LastWriteTimeUtc -Descending)
    for ($index = 0; $index -lt $files.Count; $index++) {
        $file = $files[$index]
        if ($index -lt 200 -and $file.LastWriteTimeUtc -ge $cutoff) {
            continue
        }
        $handle = $null
        $guard = $file.FullName + '.active'
        try {
            [void](Resolve-LocalPath -Path $file.FullName)
            [void](Resolve-LocalPath -Path $guard)
            Test-PrivateEntry -Path $file.FullName
            if ([IO.File]::Exists($guard)) {
                Test-PrivateEntry -Path $guard
            }
            $handle = [IO.File]::Open($guard, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            [IO.File]::Delete($file.FullName)
        }
        catch {
            # Active/inaccessible records are retained; retention is best effort and disclosed.
            $script:RetentionDeferred = $true
        }
        finally {
            if ($null -ne $handle) {
                $handle.Dispose()
                try {
                    [IO.File]::Delete($guard)
                }
                catch {
                    $script:RetentionDeferred = $true
                }
            }
        }
    }
}

function Initialize-RunRecord {
    param([string]$Directory)
    $directoryPath = New-PrivateDirectory -Path $Directory
    $name = 'outcalget-' + [datetime]::UtcNow.ToString('yyyyMMddTHHmmssZ', [Globalization.CultureInfo]::InvariantCulture) + '-' + $script:RunId + '.json'
    $path = Resolve-LocalPath -Path (Join-Path -Path $directoryPath -ChildPath $name)
    $guard = Resolve-LocalPath -Path ($path + '.active')
    $script:LogGuard = [IO.File]::Open($guard, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $script:LogGuardPath = $guard
    $script:RunLogPath = $path
    Write-RunRecord -Completed $false
    $script:RunResult.runLogPath = $path
    if ($script:DiagnosticsRequested) {
        $script:RunResult.diagnosticPath = $path
    }
    try {
        if (-not $WhatIfPreference) {
            Remove-ExpiredRunRecord -Directory $directoryPath
        }
    }
    catch {
        $script:RetentionDeferred = $true
    }
}

function Complete-RunRecord {
    # Completion evidence matters: an incomplete record means UNKNOWN, even if Scheduler shows 0.
    if ($script:AtomicCleanupFailed) {
        $script:RunResult.warnings += @{ code = 'JSON_TEMP_RETAINED'; message = 'An owned temporary JSON file could not be removed.' }
    }
    if ($script:RetentionDeferred) {
        $script:RunResult.warnings += @{
            code    = 'LOG_RETENTION_DEFERRED'
            message = 'Some old or active log files could not be pruned.'
        }
    }
    if ($script:RunResult.exitCode -eq 0 -and @($script:RunResult.warnings).Count -gt 0 -and $script:RunResult.status -ne 'preview') {
        $script:RunResult.status = 'success_with_warnings'
    }
    if ($null -ne $script:RunLogPath) {
        try {
            Write-RunRecord -Completed $true
            $script:RunResult.logCompleted = $true
        }
        catch {
            $script:RunResult.warnings += @{
                code    = 'RUN_LOG_INCOMPLETE'
                message = 'The run record could not be completed. Output may already have been published; check the published field.'
            }
            if ($script:RunResult.exitCode -eq 0) {
                $script:RunResult.exitCode = 5
                $script:RunResult.status = 'failed'
                $script:RunResult.error = @{
                    code    = 'RUN_LOG_INCOMPLETE'
                    message = 'Could not save required completion evidence.'
                    hresult = $null
                }
            }
        }
    }
    if ($null -ne $script:LogGuard) {
        $script:LogGuard.Dispose()
        try {
            [IO.File]::Delete($script:LogGuardPath)
        }
        catch {
            # An empty stale guard is harmless: retention later checks its handle, not existence.
            $script:RunResult.warnings += @{
                code    = 'LOG_GUARD_RETAINED'
                message = 'An empty run-log guard file remains.'
            }
        }
    }
    if ($script:DiagnosticsRequested -and -not $script:RunResult.logCompleted) {
        try {
            $script:RunResult.diagnostics = New-DiagnosticReport -Result $script:RunResult
        }
        catch {
            $script:RunResult.diagnostics = @{
                unavailable = $true
                code        = 'DIAGNOSTIC_UNAVAILABLE'
            }
        }
    }
}

function Read-YesNo {
    param([string]$Prompt)
    while ($true) {
        $answer = Read-Host -Prompt ($Prompt + ' [y/N]')
        if ([string]::IsNullOrWhiteSpace($answer) -or $answer -match '^(?i:n|no)$') {
            return $false
        }
        if ($answer -match '^(?i:y|yes)$') {
            return $true
        }
        Write-Information -InformationAction Continue -MessageData 'Please enter y or n.'
    }
}

function Read-CalendarSelection {
    # Local complexity exception: the three supported source choices stay together in the single interactive setup boundary.
    Write-Information -InformationAction Continue -MessageData 'Outlook Classic is ready. Setup uses its currently running Outlook profile.'
    Write-Information -InformationAction Continue -MessageData '1. Outlook folder picker (default)'
    Write-Information -InformationAction Continue -MessageData '2. Shared default Exchange calendar by owner email'
    Write-Information -InformationAction Continue -MessageData '3. Calendar already added to the Outlook Calendar navigation list'
    $choice = Read-Host -Prompt 'Source [1/2/3; Enter = 1]'
    switch ($choice) {
        { [string]::IsNullOrWhiteSpace($_) -or $_ -eq '1' } {
            return Invoke-WorkerProcess -Mode 'select' -Settings $null -StagePath $null -TimeoutSeconds 600
        }
        '2' {
            $owner = Read-Host -Prompt 'Calendar owner email address'
            Test-CalendarOwner -Owner $owner
            return Invoke-WorkerProcess -Mode 'select-shared' -Settings @{ owner = $owner } -StagePath $null -TimeoutSeconds 120
        }
        '3' {
            $data = Invoke-WorkerProcess -Mode 'nav-list' -Settings $null -StagePath $null -TimeoutSeconds 120
            $calendars = @($data.calendars)
            if ($data.unavailableCount -gt 0) {
                Write-Warning -Message 'Some navigation entries could not be opened with the current permissions; only accessible calendars are listed.'
            }
            if ($calendars.Count -eq 0) {
                Stop-Outcalget -Code 'NAVIGATION_EMPTY' -Message 'Outlook has no accessible calendars in its Calendar navigation list. Open Calendar view, add the calendar normally, or try the shared-default option.' -ExitCode 2
            }
            for ($index = 0; $index -lt $calendars.Count; $index++) {
                Write-Information -InformationAction Continue -MessageData ('{0}. {1}' -f ($index + 1), $calendars[$index].label)
            }
            $answer = Read-Host -Prompt 'Calendar number'
            $number = 0
            if (-not [int]::TryParse($answer, [ref]$number) -or $number -lt 1 -or $number -gt $calendars.Count) {
                Stop-Outcalget -Code 'SELECTION_INVALID' -Message 'Choose a calendar number from the displayed list and configure again.' -ExitCode 2
            }
            return Invoke-WorkerProcess -Mode 'select-nav' -Settings @{ fingerprint = $calendars[$number - 1].fingerprint } -StagePath $null -TimeoutSeconds 120
        }
        default {
            Stop-Outcalget -Code 'SELECTION_INVALID' -Message 'Choose source option 1, 2 or 3 and configure again.' -ExitCode 2
        }
    }
}

# Local size/complexity exception: existing transaction/rollback ordering is retained for this scheduler revision.
# Everything below is the entry point. Functions above can be loaded in isolation by tests.
if (($Help -and ($Configure -or $Check -or $Version -or $Json -or $Config -or $Worker)) -or
    ($Version -and ($Configure -or $Check -or $Json -or $Config -or $Worker))) {
    'outcalget: ARGUMENTS_INVALID. Use one operation mode; -Json applies only to export/check.'
    exit 2
}
if ($Help) {
    Get-Help -Name $PSCommandPath -Full
    exit 0
}
if ($Version) {
    '0.3.5'
    exit 0
}
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    'outcalget: HOST_LANGUAGE_BLOCKED. Existing policy prevents the required APIs. No diagnostic file can be guaranteed in this language mode.'
    exit 3
}
$script:IsWorker = [bool]$Worker
$script:ConsoleFeedback = -not $Json -and -not $Worker
$script:ConsoleFeedbackFailed = $false
$script:LastWorkerPhase = $null
$script:SyncCleanupFailed = $false
$script:SynchronizationError = $null
$script:SynchronizationConnectionObservations = @()
$script:ComCleanupWarning = $false
$script:NavigationUnavailableCount = 0
$script:OutlookStartAttempted = $false
$script:OwnedOutlookProcess = $null
$script:OutlookLease = $null
$script:LifecycleLock = $null
$script:LifecyclePending = $null
$script:LifecyclePendingOwned = $false
$script:SessionUsageLease = $null
$script:ActiveLifecycle = $null
$script:CalendarLock = $null
$script:SharedSourceLock = $null
$script:NamedExportProfile = $false
$script:ScriptPath = $PSCommandPath
$script:RunId = [Guid]::NewGuid().ToString('D')
$script:Clock = [Diagnostics.Stopwatch]::StartNew()
$script:Trace = [Collections.Generic.List[object]]::new()
$script:WorkerDiagnostics = [Collections.Generic.List[object]]::new()
$script:Phase = 'initialise'
$script:ProgressPath = $null
$script:ComAcquiredUtc = $null
$script:WorkerComAcquiredUtc = $null
$script:RemoteError = $null
$script:FailureDetails = $null
$script:SettingsSummary = @{
}
$script:UserSid = $null
$script:KeepRunArtifacts = $false
$script:StageOwned = $false
$script:StateRoot = $null
$script:RunDirectory = $null
$script:RunLogPath = $null
$script:LogGuard = $null
$script:LogGuardPath = $null
$script:RetentionDeferred = $false
$script:AtomicCleanupFailed = $false
$script:ScheduledRun = [bool]$Scheduled
$script:StartedUtc = [datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
$script:EnvironmentDiagnostics = $null
if ($script:DiagnosticsRequested) {
    try {
        $script:EnvironmentDiagnostics = Get-EnvironmentDiagnostic
    }
    catch {
        $script:EnvironmentDiagnostics = @{
            unavailable = $true
        }
    }
}
$script:Window = Get-DateWindow -Now ([datetime]::Now)
if ($Worker) {
    if ($WhatIfPreference -or $PSBoundParameters.ContainsKey('Confirm')) { exit 2 }
    $workerExit = Invoke-OutlookWorker -Directory $WorkDirectory -Token $WorkerToken
    exit $workerExit
}
$operation = 'export'
if ($Configure) {
    $operation = 'configure'
}
elseif ($Check) {
    $operation = 'check'
}
$result = [ordered]@{
    resultSchemaVersion     = 1
    version                 = '0.3.5'
    operation               = $operation
    runId                   = $script:RunId
    status                  = 'failed'
    exitCode                = 9
    published               = $false
    outputPath              = $null
    rangeStartDate          = $null
    rangeEndDate            = $null
    timeZoneId              = $null
    freshness               = 'not_verified'
    synchronization         = 'not_requested'
    internetAvailable       = $null
    outlookInitiallyRunning = $null
    outlookStarted          = $false
    outlookClosed           = $false
    joinedOwnedSession      = $false
    namedProfile            = $false
    sourceType              = 'folder'
    bytes                   = $null
    eventComponentCount     = $null
    elapsedMilliseconds     = 0
    warnings                = @()
    error                   = $null
    diagnosticPath          = $null
    diagnostics             = $null
    runLogPath              = $null
    logCompleted            = $false
    scheduled               = [bool]$Scheduled
}
if ($Check) { $result.freshness = 'not_applicable' }
$script:RunResult = $result
$configLock = $null
$outputLock = $null
$legacyOutputLock = $null
$outputClaim = $null
$stage = $null
$settings = $null
$setupCommitted = $false
$setupPublished = $false
$pending = $null
$pendingOwned = $false
try {
    if (($PSBoundParameters.ContainsKey('ExportProfile') -and $PSBoundParameters.ContainsKey('Config')) -or
        ($Configure -and $Check) -or ($Configure -and $Json) -or $WorkDirectory -or $WorkerToken) {
        Stop-Outcalget -Code 'ARGUMENTS_INVALID' -Message 'Use one operation mode. -Json applies only to export/check; worker arguments are internal.' -ExitCode 2
    }
    Add-Trace -Phase 'host-check'
    Test-HostEnvironment
    $localState = [Environment]::GetFolderPath('LocalApplicationData')
    if ([string]::IsNullOrWhiteSpace($localState)) {
        Stop-Outcalget -Code 'USER_PROFILE_UNAVAILABLE' -Message 'The current user profile has no LocalApplicationData directory. Use the configured signed-in user session.' -ExitCode 3
    }
    $script:StateRoot = Join-Path -Path $localState -ChildPath 'osp\consulthelp\outcalget'
    if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
        $LogDirectory = Join-Path -Path $script:StateRoot -ChildPath 'logs'
    }
    Add-Trace -Phase 'log-initialise'
    Initialize-RunRecord -Directory $LogDirectory
    Add-Trace -Phase 'session-check'
    Test-DesktopSession
    if ($Json -and ($WhatIfPreference -or $PSBoundParameters.ContainsKey('Confirm'))) {
        Stop-Outcalget -Code 'ARGUMENTS_INVALID' -Message 'Use preview/confirmation without -Json so PowerShell host messages cannot corrupt JSON stdout.' -ExitCode 2
    }
    $approvalParameters = @{ Debug = $false; WhatIf = [bool]$WhatIfPreference }
    if ($PSBoundParameters.ContainsKey('Confirm')) {
        $approvalParameters['Confirm'] = [bool]$PSBoundParameters['Confirm']
    }
    $approved = $Check -or (Test-ExportApproval @approvalParameters)
    if (-not $approved) {
        $result.status = 'preview'
        $result.freshness = 'not_applicable'
        $result.exitCode = 0
    }
    else {
        Add-Trace -Phase 'internet-check'
        $result.internetAvailable = $false
        Test-InternetConnection
        $result.internetAvailable = $true
        $Config = Get-ExportConfigurationPath -Name $ExportProfile -ExplicitPath $Config -NameSpecified $PSBoundParameters.ContainsKey('ExportProfile')
        $result.namedProfile = $script:NamedExportProfile
        $Config = Resolve-LocalPath -Path $Config
        if ([IO.Path]::GetExtension($Config) -ine '.json') {
            Stop-Outcalget -Code 'CONFIG_PATH' -Message 'The configuration filename must end in .json.' -ExitCode 2
        }
        if ($operation -eq 'check') {
            $result.freshness = 'not_applicable'
        }
        Add-Trace -Phase 'state-create'
        [void](New-PrivateDirectory -Path $script:StateRoot)
        $script:LockRoot = New-PrivateDirectory -Path (Join-Path -Path $script:StateRoot -ChildPath 'locks')
        [void](New-PrivateDirectory -Path (Join-Path -Path $script:StateRoot -ChildPath 'runs'))
        $script:RunDirectory = New-PrivateDirectory -Path (Join-Path -Path (Join-Path -Path $script:StateRoot -ChildPath 'runs') -ChildPath $script:RunId)
        if (-not $Check) {
            $configLock = Get-PathLock -Path (Get-CanonicalLocalPath -Path $Config) -Kind 'config'
        }
        if ($Configure) {
            if ($script:NamedExportProfile -and ([IO.File]::Exists($Config) -or [IO.Directory]::Exists($Config))) {
                Stop-Outcalget -Code 'EXPORT_PROFILE_EXISTS' -Message 'This named export profile already exists. Choose a new profile name; existing named configurations are never overwritten by setup.' -ExitCode 2
            }
            [void](Initialize-OutlookConfigurationSession)
            $selection = Read-CalendarSelection
            Write-Information -InformationAction Continue -MessageData ('Profile: ' + $selection.profileName)
            Write-Information -InformationAction Continue -MessageData ('Calendar: ' + $selection.displayPath)
            if ($selection.sourceType -ceq 'sharedDefaultCalendar') {
                Write-Information -InformationAction Continue -MessageData ('Resolved calendar owner: ' + $selection.sharedCalendarOwner)
            }
            $private = Read-YesNo -Prompt 'Include details of private appointments?'
            $attachments = Read-YesNo -Prompt 'Include attachments?'
            $binding = [Guid]::NewGuid().ToString('D')
            $defaultOutput = Join-Path -Path (Join-Path -Path (Join-Path -Path $script:StateRoot -ChildPath 'exports') -ChildPath $binding) -ChildPath 'calendar.ics'
            Write-Information -InformationAction Continue -MessageData ('Default output: ' + $defaultOutput)
            $chosen = Read-Host -Prompt 'Output path (Enter accepts the default)'
            if ([string]::IsNullOrWhiteSpace($chosen)) {
                $chosen = $defaultOutput
            }
            $newSettings = @{
                schemaVersion = 2
                bindingId = $binding
                userSid = $script:UserSid
                profileName = [string]$selection.profileName
                sourceType = [string]$selection.sourceType
                outputPath = $chosen
                includePrivateDetails = $private
                includeAttachments = $attachments
                timeoutSeconds = 120
                maxFileMiB = 25
            }
            if ($selection.sourceType -ceq 'sharedDefaultCalendar') {
                $newSettings.sharedCalendarOwner = [string]$selection.sharedCalendarOwner
                $newSettings.sharedOwnerEntryId = [string]$selection.sharedOwnerEntryId
            }
            else {
                $newSettings.storeId = [string]$selection.storeId
                $newSettings.folderEntryId = [string]$selection.folderEntryId
            }
            $settings = Read-FlatJson -Text (ConvertTo-Json -Compress -InputObject $newSettings)
            Test-Configuration -Values $settings
            if ([IO.File]::Exists($Config)) {
                Test-PrivateEntry -Path $Config
            }
            [void](New-PrivateDirectory -Path ([IO.Path]::GetDirectoryName($Config)))
        }
        else {
            Add-Trace -Phase 'config-load'
            if (-not [IO.File]::Exists($Config)) {
                Stop-Outcalget -Code 'CONFIG_MISSING' -Message 'The selected export configuration does not exist. Create it with -Configure and the same -Profile name, or use -Configure alone for the default.' -ExitCode 2
            }
            Test-PrivateEntry -Path $Config
            $settings = Read-FlatJson -Text (Read-BoundedText -Path $Config -MaxBytes 65536)
            Test-Configuration -Values $settings
        }
        if ($settings['userSid'] -cne $script:UserSid) {
            Stop-Outcalget -Code 'USER_MISMATCH' -Message 'This configuration belongs to a different Windows user. Run -Configure here.' -ExitCode 2
        }
        if ($settings['schemaVersion'] -eq 2) { $result.sourceType = $settings['sourceType'] }
        Add-Trace -Phase 'path-validate'
        $output = Resolve-LocalPath -Path $settings['outputPath']
        $result.outputPath = $output
        if ([IO.Path]::GetExtension($output) -ine '.ics') {
            Stop-Outcalget -Code 'OUTPUT_EXTENSION' -Message 'The output filename must end in .ics.' -ExitCode 5
        }
        $outputDirectory = [IO.Path]::GetDirectoryName($output)
        if ($outputDirectory.Length -gt 165) {
            Stop-Outcalget -Code 'OUTPUT_PATH_TOO_LONG' -Message 'Choose a shorter output directory (at most 165 characters) to allow a safe staging filename.' -ExitCode 5
        }
        if ($Configure) {
            [void](New-PrivateDirectory -Path $outputDirectory)
        }
        else {
            if (-not [IO.Directory]::Exists($outputDirectory)) {
                Stop-Outcalget -Code 'OUTPUT_DIRECTORY_MISSING' -Message 'The configured output directory no longer exists. Restore it or run -Configure.' -ExitCode 5
            }
            Test-PrivateEntry -Path $outputDirectory
        }
        if ([IO.File]::Exists($output)) {
            Test-PrivateEntry -Path $output
        }
        if ([IO.Directory]::Exists($output)) {
            Stop-Outcalget -Code 'DESTINATION_IS_DIRECTORY' -Message 'The output filename is occupied by a directory.' -ExitCode 5
        }
        foreach ($protectedName in @('runs', 'locks', 'profiles', 'output-claims', 'logs')) {
            $protectedPath = Get-CanonicalLocalPath -Path (Join-Path -Path $script:StateRoot -ChildPath $protectedName)
            if ((Get-CanonicalLocalPath -Path $output).StartsWith($protectedPath + '\', [StringComparison]::OrdinalIgnoreCase)) {
                Stop-Outcalget -Code 'OUTPUT_STATE_COLLISION' -Message 'Choose a calendar output outside the protected runtime, configuration and log directories.' -ExitCode 5
            }
        }
        Add-Trace -Phase 'output-claim'
        $outputClaim = Register-OutputClaim -ConfigurationPath $Config -OutputPath $output -ReadOnly ([bool]$Check)
        if ($Configure -and [IO.File]::Exists($output)) {
            Stop-Outcalget -Code 'DESTINATION_EXISTS' -Message 'Setup requires an unoccupied output filename.' -ExitCode 5
        }
        $script:SettingsSummary = @{
            schemaVersion            = $settings['schemaVersion']
            sourceType               = $result.sourceType
            namedProfile             = $result.namedProfile
            days                     = 30
            detail                   = 'full'
            includePrivateDetails    = $settings['includePrivateDetails']
            includeAttachments       = $settings['includeAttachments']
            timeoutSeconds           = $settings['timeoutSeconds']
            maxFileMiB               = $settings['maxFileMiB']
            outputExistedBefore      = [IO.File]::Exists($output)
            sourceIdentifiersPresent = $true
        }
        if ($Check) {
            $readLifecycle = Initialize-OutlookReadSession
            $checkData = Invoke-WorkerProcess -Mode 'check' -Settings $settings -StagePath $null -TimeoutSeconds ([int]$settings['timeoutSeconds']) -Lifecycle $readLifecycle
            $result.warnings = @($checkData.warnings) + @(@{
                    code    = 'EXPORT_NOT_TESTED'
                    message = 'Preflight did not test SaveAsICal, destination writes, synchronisation freshness or consumer compatibility.'
                })
        }
        else {
            $result.rangeStartDate = $script:Window.start.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            $result.rangeEndDate = $script:Window.end.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            $result.timeZoneId = [TimeZoneInfo]::Local.Id
            $outputLock = Get-PathLock -Path $outputClaim.Identity -Kind 'output'
            if (-not $outputClaim.Identity.Equals($output, [StringComparison]::OrdinalIgnoreCase)) {
                $legacyOutputLock = Get-PathLock -Path $output -Kind 'output'
                if ([IO.File]::Exists($legacyOutputLock.Name + '.pending') -or [IO.Directory]::Exists($legacyOutputLock.Name + '.pending')) {
                    Stop-Outcalget -Code 'RECOVERY_REQUIRED' -Message 'An earlier-version export left an uncertain output marker. Follow pending-marker recovery instructions before retrying.' -ExitCode 5
                }
            }
            $pending = Resolve-LocalPath -Path ($outputLock.Name + '.pending')
            if ([IO.File]::Exists($pending) -or [IO.Directory]::Exists($pending)) {
                Stop-Outcalget -Code 'RECOVERY_REQUIRED' -Message 'A previous export has an uncertain outcome. Close Outlook normally, ensure no helper is running, and follow the pending-marker recovery instructions in Get-Help before retrying.' -ExitCode 5
            }
            Write-AtomicJson -Path $pending -Value @{
                runId      = $script:RunId
                startedUtc = $script:StartedUtc
            }
            $pendingOwned = $true
            $stage = Join-Path -Path $outputDirectory -ChildPath ('.outcalget-' + $script:RunId + '.partial.ics')
            if ($Configure) {
                Write-Information -InformationAction Continue -MessageData ('Test export: Full details, ' + $result.rangeStartDate + ' through ' + $result.rangeEndDate + ', inclusive.')
            }
            $lifecycle = Initialize-OutlookSession -Settings $settings
            $exportData = Invoke-WorkerProcess -Mode 'export' -Settings $settings -StagePath $stage -TimeoutSeconds ([int]$settings['timeoutSeconds']) -Lifecycle $lifecycle
            $result.warnings = @($exportData.warnings)
            Add-Trace -Phase 'file-validate'
            [void](Resolve-LocalPath -Path $stage)
            Test-PrivateEntry -Path $stage
            [void](Get-CanonicalLocalPath -Path $stage)
            $validation = Test-IcsFile -Path $stage -MaxFileMiB ([int]$settings['maxFileMiB'])
            Test-InternetConnection
            Add-Trace -Phase 'publish'
            [void](Resolve-LocalPath -Path $output)
            Test-PrivateEntry -Path $outputDirectory
            if ([IO.File]::Exists($output)) {
                Test-PrivateEntry -Path $output
            }
            if ((Get-CanonicalLocalPath -Path $output) -cne $outputClaim.Identity) {
                Stop-Outcalget -Code 'OUTPUT_IDENTITY_CHANGED' -Message 'The destination identity changed during export. The previous output was not replaced.' -ExitCode 5
            }
            Publish-CalendarFile -Stage $stage -Destination $output -MustBeNew ([bool]$Configure)
            $result.published = $true
            $result.bytes = $validation.bytes
            $result.eventComponentCount = $validation.eventComponentCount
            if ($Configure) {
                $setupPublished = $true
                Add-Trace -Phase 'config-save'
                Write-AtomicJson -Path $Config -Value $settings
                $setupCommitted = $true
            }
        }
        Add-Trace -Phase 'complete'
        $result.status = 'success'
        if (@($result.warnings).Count -gt 0) {
            $result.status = 'success_with_warnings'
        }
        $result.exitCode = 0
    }
}
catch {
    $failure = Get-ExceptionInfo -Record $_
    if ($null -ne $script:RemoteError) {
        $failure.hresult = $script:RemoteError.hresult
        $failure.exceptionTypes = @($script:RemoteError.exceptionTypes)
        $failure.line = [int]$script:RemoteError.line
        if ($script:RemoteError.PSObject.Properties.Name -contains 'phase') {
            $failure.phase = [string]$script:RemoteError.phase
        }
    }
    $script:FailureDetails = @{
        exceptionTypes = $failure.exceptionTypes
        scriptLine     = $failure.line
        hresult        = $failure.hresult
        phase          = $failure.phase
    }
    $result.error = @{
        code    = $failure.code
        message = $failure.message
        hresult = $failure.hresult
    }
    $result.exitCode = $failure.exitCode
    if ($failure.exitCode -eq 8) {
        $result.status = 'cancelled'
    }
}
finally {
    try {
        Complete-OutlookSession -Settings $settings
    }
    catch {
        $script:KeepRunArtifacts = $true
        $cleanup = Get-ExceptionInfo -Record $_
        if ($null -eq $script:FailureDetails) {
            $script:FailureDetails = @{ exceptionTypes = $cleanup.exceptionTypes; scriptLine = $cleanup.line; hresult = $cleanup.hresult; phase = $cleanup.phase }
        }
        $result.warnings += @{ code = $cleanup.code; message = $cleanup.message }
        if ($result.exitCode -eq 0) {
            $result.status = 'failed'
            $result.exitCode = 5
            $result.error = @{ code = $cleanup.code; message = $cleanup.message; hresult = $cleanup.hresult }
        }
    }
    if ($script:SyncCleanupFailed) {
        $result.warnings += @{ code = 'SYNC_CLEANUP_FAILED'; message = 'A synchronization setting or event subscription could not be restored. Inspect Outlook and the pending markers before another run.' }
        $script:KeepRunArtifacts = $true
    }
    if ($script:ComCleanupWarning) {
        $result.warnings += @{ code = 'COM_RELEASE_WARNING'; message = 'A native object reference did not release cleanly.' }
    }
    if ($script:ConsoleFeedbackFailed) {
        $result.warnings += @{ code = 'CONSOLE_FEEDBACK_UNAVAILABLE'; message = 'Some console feedback could not be displayed; use the run record.' }
    }
    if ($Configure -and $setupPublished -and -not $setupCommitted) {
        try {
            [IO.File]::Delete($result.outputPath)
            $result.published = $false
            $result.bytes = $null
            $result.eventComponentCount = $null
        }
        catch {
            $result.warnings += @{
                code    = 'SETUP_OUTPUT_RETAINED'
                message = 'The new setup output remains, but configuration was not saved.'
            }
        }
    }
    if ($null -ne $stage -and [IO.File]::Exists($stage) -and $script:StageOwned -and -not $script:KeepRunArtifacts) {
        try {
            [IO.File]::Delete($stage)
        }
        catch {
            $result.warnings += @{
                code    = 'STAGING_RETAINED'
                message = 'A protected staging file could not be removed.'
            }
        }
    }
    elseif ($null -ne $stage -and [IO.File]::Exists($stage) -and -not $script:StageOwned -and -not $script:KeepRunArtifacts) {
        $result.warnings += @{
            code    = 'STAGING_COLLISION_RETAINED'
            message = 'An unexpected staging file was left untouched because this run did not own it.'
        }
    }
    if ($pendingOwned -and -not $script:KeepRunArtifacts) {
        try {
            [IO.File]::Delete($pending)
        }
        catch {
            $result.warnings += @{
                code    = 'PENDING_MARKER_RETAINED'
                message = 'A recovery marker remains; subsequent exports will stop until recovery.'
            }
            if ($result.exitCode -eq 0) {
                $result.status = 'failed'
                $result.exitCode = 5
                $result.error = @{
                    code    = 'PENDING_MARKER_RETAINED'
                    message = 'The export was published, but its recovery marker could not be cleared. Follow the recovery instructions before another export.'
                    hresult = $null
                }
            }
        }
    }
    if ($Configure -and $null -ne $outputClaim -and $outputClaim.Created -and -not $setupCommitted -and -not $script:KeepRunArtifacts) {
        try {
            $catalogLock = Wait-PathLock -Path (Join-Path -Path $script:StateRoot -ChildPath 'output-catalog') -Kind 'catalog' -TimeoutSeconds $CoordinationTimeoutSeconds
            try { [IO.File]::Delete($outputClaim.Path) }
            finally { $catalogLock.Dispose() }
        }
        catch {
            $result.warnings += @{ code = 'OUTPUT_CLAIM_RETAINED'; message = 'A reservation from failed setup remains. Review output ownership before reusing that filename.' }
        }
    }
    if ($null -ne $script:OwnedOutlookProcess) { $script:OwnedOutlookProcess.Dispose() }
    if ($null -ne $script:LifecycleLock) { $script:LifecycleLock.Dispose() }
    if ($null -ne $script:CalendarLock) { $script:CalendarLock.Dispose() }
    if ($null -ne $script:SharedSourceLock) { $script:SharedSourceLock.Dispose() }
    if ($null -ne $legacyOutputLock) { $legacyOutputLock.Dispose() }
    if ($null -ne $outputLock) {
        $outputLock.Dispose()
    }
    if ($null -ne $configLock) {
        $configLock.Dispose()
    }
    $result.elapsedMilliseconds = $script:Clock.ElapsedMilliseconds
    if ($null -ne $script:RunDirectory -and -not $script:KeepRunArtifacts) {
        try {
            [IO.Directory]::Delete($script:RunDirectory, $true)
        }
        catch {
            $result.warnings += @{
                code    = 'RUN_FILES_RETAINED'
                message = 'Protected operational files could not be removed. Do not share the runs directory.'
            }
        }
    }
    if ($script:KeepRunArtifacts) {
        $result.warnings += @{
            code    = 'RUN_ARTIFACTS_RETAINED'
            message = 'Protected run/staging files were retained because Outlook may still be writing. Do not share them.'
        }
    }
    if ($result.exitCode -eq 0 -and @($result.warnings).Count -gt 0 -and $result.status -ne 'preview') {
        $result.status = 'success_with_warnings'
    }
    Complete-RunRecord
    if ($result.exitCode -eq 0 -and @($result.warnings).Count -gt 0 -and $result.status -ne 'preview') {
        $result.status = 'success_with_warnings'
    }
}
if ($Json) {
    ConvertTo-Json -InputObject $result -Depth 10 -Compress
}
else {
    if ($result.exitCode -eq 0) {
        if ($result.status -eq 'preview') {
            Write-Output -InputObject 'Preview: no calendar export or configuration changes were performed. Run logging may have occurred.'
        }
        elseif ($Check) {
            Write-Output -InputObject 'Preflight passed. Native export access and write capability remain untested.'
        }
        else {
            $syncSummary = Get-SynchronizationSummary -Synchronization $result.synchronization
            Write-Output -InputObject ('Exported {0} VEVENT components to {1}. Range {2} through {3}. {4}' -f $result.eventComponentCount, $result.outputPath, $result.rangeStartDate, $result.rangeEndDate, $syncSummary)
        }
    }
    else {
        $failurePhase = $script:Phase
        if ($null -ne $script:FailureDetails -and $script:FailureDetails.ContainsKey('phase')) {
            $failurePhase = $script:FailureDetails.phase
        }
        [Console]::Error.WriteLine(('outcalget failed [{0}] in phase {1}: {2}' -f $result.error.code, $failurePhase, $result.error.message))
        if (-not $result.published) {
            Write-Output -InputObject 'No final export was refreshed by this run.'
        }
        else {
            Write-Output -InputObject 'An export was published before the later failure; inspect the output before retrying.'
        }
    }
    foreach ($warning in @($result.warnings)) {
        Write-Warning -Message ('[{0}]: {1}' -f $warning.code, $warning.message)
    }
    if ($null -ne $result.runLogPath) {
        Write-Information -InformationAction Continue -MessageData ('Run record: ' + $result.runLogPath)
    }
    if ($null -ne $result.diagnostics) {
        Write-Output -InputObject 'Sanitised diagnostic report (copy this JSON):'
        ConvertTo-Json -InputObject $result.diagnostics -Depth 12
    }
}
exit ([int]$result.exitCode)
