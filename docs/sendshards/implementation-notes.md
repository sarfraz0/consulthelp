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
