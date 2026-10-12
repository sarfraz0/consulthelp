# outcalget usage

[Documentation index](README.md)

## Parameters

### -Configure

Interactive source/options selection and a test export. Saves only after success and
requires an unoccupied output filename. Never put this switch in a scheduled action.
With -Profile, creates a new named configuration and refuses an existing named file.
Starts Classic if closed, waits for its profile/stores and online Exchange, then offers
selection. The selected calendar is synchronized before the test export when the
session is script-owned. Existing user sessions are left open and receive no explicit
sync. Only this run's owned startup process is eligible for normal Quit during cleanup.

### -ExportProfile

Alias: -Profile. Names an export configuration; no selector or default reuses the
original config.json. This never selects or changes the Outlook MAPI profile.

### -Config

Legacy explicit absolute local .json path; mutually exclusive with -Profile.

### -Check

Read-only business preflight of config/source/apparent path/output ownership. Requires
already-open Outlook. No native export, output write, claim registration, explicit sync
or freshness/consumer test. Operational logs, locks and private worker files can occur.

### -Scheduled

Non-prompting saved export plus scheduled metadata. Does not create a task or alter
business behavior. Incompatible with Configure, Check, Help and Version parameter sets.

### -CoordinationTimeoutSeconds

Default 300 seconds, range 15-1800. Applies to each catalog/lifecycle/COM-queue or
joined-session cleanup wait. Worker queue time is additional to its native-work budget;
parent monitoring detects native start separately and also enforces a total deadline.
A busy same-configuration/calendar/output lock fails immediately rather than queuing.

### -StartupTimeoutSeconds

Export-only readiness budget, default 120 seconds, range 15-600. Includes native COM,
saved-source and primary online Exchange readiness for script-owned sessions; blocking
COM work is isolated in a bounded worker process. Configure uses the embedded defaults.

### -SyncTimeoutSeconds

Export-only native group synchronization budget, default 180 seconds, range 30-600.
Applies to a process started here and to joined script-owned sessions. Existing user
sessions receive no explicit sync. Native sync completion is not proof of freshness.
Connection recovery shares this same budget, never a fresh timeout. At most three
native starts are allowed; each interrupted attempt is stopped/detached and a new
correlated completion is required after recovery. Explicit Work Offline fails promptly.

### -LogDirectory

Absolute approved private local directory; default state-root\logs. Does not relocate
configuration, ownership claims or operational state. Overrides the embedded log path.

### -Json

One result JSON envelope on stdout, including handled failures. Schema 1 retains all
existing fields and adds namedProfile, sourceType and joinedOwnedSession. Debug does
not change it. Do not combine with WhatIf/Confirm, whose host output can corrupt JSON.

### -Help

Compatibility alias for full Get-Help; also supports Get-Help and -?.

### -Version

Prints 0.3.5 without Outlook access or run records.

## Examples

### Example 1

```powershell
powershell.exe -NoLogo -NoProfile -STA -File "C:\ApprovedTools\outcalget.ps1" -Configure -Profile team -Debug
```

Classic starts automatically if closed. Choose option 3 for a calendar already in the Calendar
navigation list, or option 2 for a shared mailbox's default calendar. Choose a distinct
output path. Later runs do not ask for a calendar, owner, options or output path.

### Example 2

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\ApprovedTools\outcalget.ps1" -Scheduled -Debug
```

Exports the existing default configuration without reconfiguration.

### Example 3

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\ApprovedTools\outcalget.ps1" -Scheduled -Profile team -Debug
```

Headless named export. Use the same arguments in that profile's Task Scheduler action.

### Example 4

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\ApprovedTools\outcalget.ps1" -Check -Profile team -Json
```

Read-only preflight. Outlook must already be open. Does not replace a real export test.

### Example 5

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -File "C:\ApprovedTools\outcalget.ps1" -Scheduled -Profile team -WhatIf -Debug
```

Preview: no source access, export, startup, output claim or configuration changes.
Minimal run logging may occur; retention is skipped. Configuration/source is not validated.

## Inputs

None. Configurations are generated JSON data. Private worker switches are not public API.

## Outputs

Normal CLI: UTC phase feedback, export/check summary, sanitized error on stderr, warnings
and run-record path. Native hosts can render information/warnings on stdout. -Json
suppresses phase feedback and provides one result envelope. Nonzero exits distinguish
failure. Output ICS bytes are native Outlook output, not manually constructed events.

## Exit codes

0 success/check/preview/help/version; 2 configuration/binding/output ownership;
3 host/session/Internet/Outlook/access; 4 native export; 5 storage/logging/recovery;
6 unsafe overlap; 7 bounded timeout/coordination; 8 handled cancellation; 9 protocol/
unexpected error. Pre-body and abrupt-stop failures can use host/Scheduler codes.
