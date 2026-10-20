# sendshards implementation notes

[Documentation index](README.md)

These notes preserve function help and explanatory comments removed from the script. Original source line numbers refer to the pre-extraction file. The extraction preserved executable tokens and strings; the implementation changes recorded below apply to revision 1.0.1.

## ConvertTo-NativeArgument

Original line 281: Windows CRT quoting: double slashes before a quote and before the closing quote.

## Stop-OwnedProcess

Original line 330: Never kill by image name. The PID belongs to the Process object started here.

## Send-ItemCommit

Original line 1473: Exact source OID, one branch refspec, no force, no tag propagation, hooks active.

## Invoke-TransferPush

Original line 1510: Save the original start before contacting the remote or making Git mutations.

## Invoke-PlanPublication

Original line 1617: Revalidate the frozen view after acquiring the tool lock.

## Entry point

Original line 1633: Thin executable boundary: no dot-sourcing interface, success objects, or custom prompts.

Original line 1650: JSON reader assembly is loaded here before reading the input document.

## Revision 1.0.1 corrections

- `Invoke-TransferGit` explicitly selects floating-point `Math.Min`, retaining the integer CLI timeout while allowing an initially infinite planning budget and fractional remaining time.
- `Write-AtomicJson` passes `[System.Management.Automation.Language.NullString]::Value` directly to `File.Replace`, preserving same-filesystem atomic replacement without an empty backup filename. The type is available in Windows PowerShell 5.1. Existing-state and NewOnly protection remain.
- `Get-JsonPropertyName` uses the native JSON reader’s decoded `item` attribute where present, otherwise `LocalName`. Duplicate detection is ordinal and case-insensitive at every object level. This fixes escaped-key collisions and distinct non-XML-name false collisions without a custom JSON parser. Native `__type` metadata is rejected because no tool schema permits it and the reader can map its first occurrence to an attribute.
- `Read-TransferJson` retains root string values from the bounded native reader and restores those strings after `ConvertFrom-Json`. All persisted timestamp fields are root properties. This preserves their exact decoded input spelling across 5.1 and Core, including rejecting noncanonical timestamps later; it does not stringify arrays, numbers or objects. No PowerShell-7-only switch or runtime module is required.
- `Assert-SourceManifest`, `Assert-ManifestMatch` and `Read-TransferState` validate the missing string types/formats before identity comparisons, including checkpoint PlanId/PlanHash/Status.
- `Invoke-TransferPush` requires an empty index when all pending commits already exist, including zero-pending and completed resumes, and checks again before reporting completion. An exact staged next shard remains resumable.

Analyzer corrections rename four private helpers to singular nouns (`Get-SourceRecord`, `Assert-ShardGroup`, `Assert-CanonicalPlanPath`, `Assert-PlanRecord`), rename the event parameter while retaining its `Event` alias, deliberately discard asynchronous cleanup-drain return tasks, and correct `Get-TransferChain`’s output declaration to its existing object-array return. Public CLI and plan/state schemas are unchanged.

Seven function-scoped `PSUseShouldProcessForStateChangingFunctions` exceptions are explicit analyzer configuration, not a general suppression: `Stop-Transfer` only throws; `New-TransferContext` constructs context/checks prerequisites; `Set-RepositoryContext` reads local metadata; `New-TransferPlan` builds a local in-memory plan; `Set-TransferDeadline` updates in-memory timing; `New-ItemCommit` runs inside the script’s ShouldProcess-approved Push; `Stop-OwnedProcess` must clean up the exact child on cancellation without a new prompt. The script’s public ShouldProcess boundary remains active.

Native mapping reference: [Microsoft JSON/XML mapping](https://learn.microsoft.com/en-us/dotnet/framework/wcf/feature-details/mapping-between-json-and-xml). File-replacement and PowerShell-version references remain in the [original audit](audit-2026-10-18.md).

## Revision 1.0.2 corrections

`Read-TransferState` refuses a missing checkpoint when the original start record survives. `Assert-TransferStart` validates its strict schema/identity/original timestamp; `Write-TransferStart` creates it once with the existing bounded atomic writer. `Invoke-TransferPush` also refuses missing state with exact local transfer commits, covering older plans without an anchor. A valid older checkpoint is upgraded without changing its start. `Assert-ExternalPlanPath` and `Invoke-PlanPublication` protect the additional runtime data path. The plan and checkpoint schemas remain version 1; the separate four-field start record has SchemaVersion 1. Recovery/custody limitations are explicit in operations.

`Initialize-OwnedGitType` embeds a Windows 10+ Job Object adapter. `Start-OwnedGit` uses it for every Git invocation; `Stop-OwnedProcess`/taskkill have been removed. CreateProcess uses CREATE_SUSPENDED, CREATE_UNICODE_ENVIRONMENT, EXTENDED_STARTUPINFO_PRESENT and CREATE_NO_WINDOW, with explicit std-handle and job lists. Assignment at process creation closes the hard-termination window between spawning and later job assignment. No breakaway is requested. The job handle is not inherited; SafeHandle/finally/OS closure keeps kill-on-close ownership after parent exit. Bounded stream reads, timeouts, inherited environment and Git controls remain in the PowerShell adapter. Cleanup checks job active-process accounting, retains primary failure classification, and fails an otherwise successful call when cleanup cannot be confirmed.

The six unchanged helper ShouldProcess exceptions from 1.0.1 remain; the removed `Stop-OwnedProcess` exception is replaced by a function-scoped `Start-OwnedGit` adapter exception. The entry point still gates publication/mutations; local planning queries are read-only Git invocations. Cleanup cannot prompt or be declined independently.

### Function-named engineering exceptions

Current AST counting uses the same documented clause/loop/catch/Boolean, container-nesting and token-body convention as the companion audit. Existing focused exceptions remain: `Read-TransferJson` complexity 12/nesting 4 for bounded native-reader structural and root-string checks; `Test-JsonElement` complexity 11 for strict primitive/schema validation; `Assert-ControlledDirectory` nesting 4 for ancestor/type/link checks. No blanket analyzer/metric suppression is added.

`Initialize-OwnedGitType` has 217 body lines when the embedded C# source is counted. This is a deliberate single-file native-boundary exception: declarations, explicit handle inheritance, atomic job association, safe handle closure and cancellation cleanup must ship together without a runtime module or launcher. The C# Start method also exceeds the default 80-line body guidance; separating allocation from its finally resource ledger would obscure startup failure ownership. C# compilation/ABI checks are development evidence, not Windows API execution proof.

Vendor references: [Job Objects](https://learn.microsoft.com/en-us/windows/win32/procthread/job-objects), [UpdateProcThreadAttribute: HANDLE_LIST/JOB_LIST](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-updateprocthreadattribute), [InitializeProcThreadAttributeList](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-initializeprocthreadattributelist), [redirected child-process pipes](https://learn.microsoft.com/en-us/windows/win32/procthread/creating-a-child-process-with-redirected-input-and-output), [job accounting](https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-jobobject_basic_accounting_information). No policy bypass or unsupported fallback is included.
