# outcalget implementation notes

[Documentation index](README.md)

These notes preserve the comments removed from `outcalget.ps1`, grouped by function. Original line numbers identify the pre-extraction source only. Analyzer attributes and their justifications remain in the implementation because they configure analysis. Native C# bridge source strings remain byte-for-byte unchanged.

## Script initialization

Original line 494: -Debug is PowerShell's common parameter. Consume it without Write-Debug/Inquire prompts.

## Stop-Outcalget

Original line 508: Optional phase metadata can be absent in a helper-only host before entry initialization.

## Get-ExceptionInfo

Original line 530: Only messages constructed by Stop-Outcalget.

## Write-AtomicJson

Original line 615: Preserve the primary write/publication error; report cleanup at the boundary.

## Read-FlatJson

Original line 646: Local complexity exception: preserve the tested flat-schema duplicate-key validation; native JSON still decodes values.

Original line 648: Tokenise the deliberately flat scalar schema before using the built-in JSON decoder.

Original line 649: This detects duplicate/escaped duplicate keys that ConvertFrom-Json can otherwise discard.

## Test-Configuration

Original line 710: Local complexity exception: exact conditional schemas preserve legacy bindings and fail closed on unknown source types.

## Test-PrivateEntry

Original line 887: Conservative: do not guess whether an unfamiliar corporate group is privileged.

## Initialize-PathIdentityBridge

Original line 970: Read-only metadata: volume-GUID paths detect drive aliases; link count rejects hard-linked targets.

Original line 971: OPEN_EXISTING + zero access never creates/truncates a file. No privilege or policy changes.

## Test-IcsLine

Original line 1140: Local complexity exception: keep component state validation in one place to preserve native output acceptance.

## Test-IcsFile

Original line 1200: Local size/complexity exception: bounded reader, unfolding and disposal belong to one validation lifetime.

## Test-ExportApproval

Original line 1304: A bound -Debug can force ShouldProcess confirmation independently of DebugPreference.

Original line 1305: This adapter is called with -Debug:$false; explicit Confirm/WhatIf are forwarded.

## Test-DesktopSession

Original line 1310: Interactive-token tasks share the user's desktop. Password/S4U tasks do not.

## Write-PhaseMessage

Original line 1359: Console feedback is optional; the required sanitized run record remains authoritative.

## Test-InternetConnection

Original line 1366: Local complexity exception: fail-closed status/type/error classification and native-reference cleanup form one check.

## Get-OutlookExecutable

Original line 1429: Only machine-controlled App Paths are used; never PATH, a shell association or HKCU.

## Start-OutlookSession

Original line 1460: Interactive setup has no saved binding yet: use normal Outlook startup.

## Get-ReadyOutlookSetup

Original line 1520: Local complexity exception: one bounded poll owns profile/store/connection checks and the per-attempt Stores reference lifetime.

## Get-ReadyOutlookSource

Original line 1552: A newly started profile can publish its COM application before stores load.

## Get-SharedCalendarBinding

Original line 1587: Local complexity exception: recipient resolution, mailbox identity pinning and transferred folder ownership share one COM-reference lifetime.

Original line 1600: This binding targets an Exchange directory mailbox, rather than an ambiguous display name or one-off recipient.

Original line 1611: Entry IDs can have different binary representations for the same object.

Original line 1612: Ask MAPI to compare the pinned owner rather than compare the strings.

## Get-NavigationCalendar

Original line 1695: Local complexity exception: indexed COM traversal owns and releases each navigation reference in one scope.

## Initialize-SynchronizationBridge

Original line 1773: A narrow native-event bridge avoids an optional Outlook PIA dependency and

Original line 1774: PowerShell callbacks on COM threads. GUID/DISPIDs verified against Microsoft's

Original line 1775: Outlook interop metadata: SyncStart 61441, OnError 61443, SyncEnd 61444.

## Wait-CalendarSynchronization

Original line 1886: Local complexity exception: bounded connection recovery, correlated events and cleanup share one synchronization lifetime and deadline.

Original line 1913: Completion from a disconnected attempt is never reused after recovery.

## ConvertTo-SafeConnectionObservation

Original line 1983: Copy only numeric/boolean metadata, never extra fields from the reply.

## Invoke-CalendarSynchronization

Original line 1994: Local complexity exception: native setting checkpoint/restoration and RCW release share one failure-safe ownership transaction.

Original line 2004: Microsoft documents Store=null for GetSharedDefaultFolder. The caller

Original line 2005: permits this only after revalidating the pinned Exchange owner binding.

Original line 2029: Capture the primary native stage before restoration records cleanup phases.

## Close-OutlookSession

Original line 2080: Local complexity exception: every native user-activity guard must complete before the single irreversible Quit call.

## Read-WorkerFeedback

Original line 2127: Local complexity exception: the correlated allowlist and optional native-start timestamp share one bounded progress read.

Original line 2142: Native JSON decoders can materialize ISO timestamps as DateTime.

## Get-OutlookCoordinatorState

Original line 2167: Local complexity exception: exact marker shape and live PID/start-time/session checks fail closed as one ownership decision.

## Open-OutlookUsageLease

Original line 2237: The lifecycle gate prevents another parent from racing initial file creation.

## Enter-OutlookCoordination

Original line 2252: Local complexity exception: the short gate makes process ownership, marker validation and shared-lease acquisition indivisible between cooperating parents.

## Complete-OutlookSession

Original line 2399: Local complexity exception: owner quiescence, reference leases, normal Quit and uncertain-state preservation share a single cleanup lifetime.

## Invoke-OutlookWorker

Original line 2481: Local size/complexity exception: preserve the working COM acquisition, export, reply and release transaction.

Original line 2545: A terminated COM worker can leave Outlook still working. Never start another

Original line 2546: native operation behind an uncertain session marker, even for another output.

Original line 2725: The stage was vacant when this worker accepted it. Only a worker that

Original line 2726: reached the native write may authorise parent cleanup of this pathname.

## Invoke-WorkerProcess

Original line 2773: Local size/complexity exception: process, timeout, reply validation and disposal share one ownership lifetime.

Original line 2818: Drain to a null stream: never retain or forward raw host/COM error text.

Original line 2863: Only explicitly safe diagnostics are copied; source selection data stays private.

## New-RunRecord

Original line 3113: Optional enrichment cannot prevent required minimal operational logging.

## Write-RunRecord

Original line 3128: Keep a bounded diagnostic trace; never lose the minimal outcome to optional detail.

## Remove-ExpiredRunRecord

Original line 3146: The per-run .active handle prevents deletion of any still-running parent's record.

Original line 3147: Only this exact tool-owned name format is eligible. Sensitive runs/staging are NEVER pruned here.

Original line 3172: Active/inaccessible records are retained; retention is best effort and disclosed.

## Complete-RunRecord

Original line 3214: Completion evidence matters: an incomplete record means UNKNOWN, even if Scheduler shows 0.

Original line 3254: An empty stale guard is harmless: retention later checks its handle, not existence.

## Read-CalendarSelection

Original line 3289: Local complexity exception: the three supported source choices stay together in the single interactive setup boundary.

## Entry point

Original line 3329: Local size/complexity exception: existing transaction/rollback ordering is retained for this scheduler revision.

Original line 3330: Everything below is the entry point. Functions above can be loaded in isolation by tests.

## update-calendar.ps1 feedback correction, 20 October 2026

The wrapper still uses the exporter’s JSON result as the publication contract. `Read-CalendarExportResult` parses it once, preserves process failure, and displays the exporter’s allowlisted error code/message rather than discarding the failure envelope. `Get-CalendarExportMessage` shares scalar-type checks for warning/error display. `Write-CalendarExportNotice` reports the optional run-record location and structured warnings on information/warning streams, leaving the success stream available for the returned object. Malformed/raw stdout is not echoed. A published-before-failure result is disclosed and still stops Git publication.

The entry point announces launch before waiting for the child, confirms completion only after `Get-PublishedCalendarPath` validates the successful publication, and explicitly identifies a skipped ShouldProcess operation. Individual Outlook phases remain suppressed by `-Json`. `-Verbose` identifies the resolved adjacent or explicit exporter. Exporter/COM/worker behavior, configuration selection, previous-output safeguards and Git path/location semantics are unchanged.

## update-calendar.ps1 scheduling audit, 21 October 2026

`Get-CalendarPushTarget`, `Get-CalendarRemoteTip` and `Assert-CalendarHistory` replace an unrestricted bare push with the current branch's approved upstream remote/branch. A push is necessary whenever HEAD differs from the inspected remote tip, including an unchanged export after an earlier rejected push. Pending linear history is checked per commit for this literal calendar only, so unrelated unpublished content cannot ride along. Tags, submodule pushes and different/multiple explicit push refspecs and conflicting push-remote overrides are refused. `Get-CalendarPushTarget` retains a local complexity exception to keep these destination consistency checks together. Remote names remain intact for pre-push hooks and remote-specific transport/proxy/credential controls. For a separate push URL, the read-only query clears and sets only the command-local remote URL list to inspect that endpoint; it does not write configuration or weaken authentication/TLS. The destination is rechecked before push. `Assert-CalendarRepositoryReady` prevents staging from resolving an existing merge/conflict accidentally.

`Invoke-CalendarGit` uses a directly started process, .NET Framework-supported concurrent ReadToEndAsync streams and bounded process/pipe waits, with Windows native argument quoting in `ConvertTo-CalendarArgument`. Git/GCM prompting is disabled only in the child environment; signing, SSH and custom hooks remain provisioned native dependencies. Output is buffered in memory and discarded rather than logged. This is suitable for reviewed controlled-session helpers, not a hostile-output memory sandbox. A narrow size/complexity exception keeps process, marker, timeout, stream completion and disposal in one ownership lifetime. Timeout can stop only the exact Git process and leaves native descendants uncertain; a per-user pending marker blocks the next update until manual review. Normal exit, including hook/push refusals, clears the marker. Marker-deletion failure warns without replacing a primary error and is recorded as requiring recovery. `Stop-CalendarUpdate` raises only a classified exception; its narrow analyzer attribute does not suppress mutation checks elsewhere.

`Initialize-CalendarRecord` and `Write-CalendarRecord` provide a single allowlisted JSON report, before business work and at phase boundaries. The open write handle prevents cooperating retention from deleting active records. Retention/listing and optional debug hashing failures warn; required record writes fail the workflow. Normal and debug records contain no raw native messages, paths, profile/remote names or calendar content. A local complexity exception keeps initialization, optional enrichment and best-effort retention together; the report is rewritten in place, so a mid-write read can be temporarily unreadable. Missing/stale/incomplete evidence means unknown. An executable entry-block size/complexity exception keeps ShouldProcess, the exclusive per-user wrapper lock, export/result validation, single-file commit, verified push and final record/lock cleanup under one boundary. Handled completion is distinct from finally execution, so cancellation cannot be promoted to a completed success record. The existing `Read-CalendarExportResult` scalar validation/error branches retain their local complexity exception; only a valid GUID and typed publication/exit fields are copied into wrapper logs.

The exporter remains byte-identical at 0.3.5; its native COM budgets, output/config/source locks, publication, synchronization restoration, Outlook ownership, privacy rules and recovery contracts are unchanged. New wrapper Scheduled/Debug forwarding adds metadata/diagnostics without changing those business choices. Documentation/XML remain optional; the existing adjacent exporter is the only companion implementation dependency. [Verification](verification.md#scheduled-update-audit-21-october-2026) separates synthetic local evidence from pending Windows acceptance.

## Date-range and startup-window follow-up, 22 October 2026

`Get-DateWindow` already returns local Date through Date.AddDays(29); there is no arithmetic correction. `Set-CalendarExportWindow` keeps those native settings together and verifies their read-back before saving, without interpreting or rewriting native ICS recurrence data. The worker still validates its correlated ISO date request before constructing CalendarSharing and retains full-detail/privacy/attachment/working-hours settings. `New-RunRecord` now includes the already-public requested range/timezone and startup outcome, making normal scheduled records useful for this report. An old DTSTART plus RRULE is valid recurrence representation, not evidence that the date request was widened; actual exported recurrence/consumer behavior needs target inspection.

`Set-OutlookStartupWindow` uses the documented Outlook Explorer.WindowState property and OlWindowState olMinimized=1, once, inside the existing bounded COM worker. It supplements the existing ProcessStartInfo minimized launch hint, which Outlook can override. The parent requests it only for Mode=ready, ScheduledRun=true and RunResult.outlookStarted=true; shared/user-open/configuration/later workers never request it. The worker validates the private Boolean/mode/lease request and the parent allowlists the status returned. The helper checks the exact ownership lease before inspecting and again before changing one normal Explorer, skips user-activity evidence and releases all acquired references through the existing Close-ComReference adapter. Its small local complexity exception keeps ownership/user-activity gates, optional native outcome and COM cleanup in one lifetime. Optional minimization failure is a fixed status/warning, not a raw exception or permission workaround; classified ownership errors still fail. Readiness warnings are now retained when export warnings are appended. The code neither creates windows nor repeats minimization later to fight a user restore. No invisible-startup guarantee is made for Outlook/Office/add-in dialogs.

The APIs above are native Outlook interfaces; their reference links are listed in the README. No public-document network lookup or real Outlook access was performed during this restricted local review. New parameters, modules, setup, elevation, policy changes, recurrence expansion and custom ICS rewriting are outside this change. Existing startup/close ownership, source/privacy/coordination, staging/recovery and standalone packaging remain intact. See the current [verification](verification.md#date-window-and-startup-visibility-22-october-2026) for fixture and target boundaries.
