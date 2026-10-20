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
