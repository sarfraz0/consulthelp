# outcalget operations

[Documentation index](README.md)

## Required Task Scheduler Settings

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

## Configuration And Concurrency

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

## Run Records

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

## Recovery

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

## Synchronization And Closure Limits

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

## Security / Compatibility

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

## Upgrade And Acceptance

Replace the script and matching external help while no helper is running; existing configuration and scheduled
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

## Functional Acceptance For A Shared Calendar

A completed synchronization group or exit 0 alone does not prove server freshness
or event completeness. After setup, inspect the actual generated .ics for a known
appointment unique to that shared calendar within the next 30 local dates, checking
its date/time and details. Compare with a manual full-details Outlook export using
the same window; include recurring and all-day appointments when present. A valid
empty file is allowed for a genuinely empty window, so check expected content rather
than treating a nonzero event count as a source guarantee. Development fixtures
exercise source routing and publication, not real Exchange event data.
