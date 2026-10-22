# embelish implementation notes

[Documentation index](README.md)

These notes preserve function help and explanatory comments removed from the script. Original source line numbers refer to the pre-extraction file. The extraction preserved executable tokens and strings. Revision 1.1.1 additionally resolves an explicit diagnostic directory through PowerShell’s filesystem path resolver before passing it to .NET file I/O; this fixes relative log paths when PowerShell and process working directories differ.

## Script initialization

Original line 183: Run-scoped constants and diagnostic state. No secret is ever stored here.

## Get-EmbelishTimestamp

### Synopsis

Returns the current UTC time as an ISO-8601 string.

## Write-DiagnosticEvent

### Synopsis

Appends one sanitised JSON Lines diagnostic record to the run log.

### Description

Logging failures never abort the business operation; they degrade to a
single warning. Callers must pass only non-secret fields in -Data.

## Initialize-DiagnosticContext

### Synopsis

Resolves the log directory, opens a per-run log file and prunes old logs.

## Resolve-SevenZipPath

### Synopsis

Locates the 7-Zip executable, honouring an explicit path first.

## Invoke-SevenZipCommand

### Synopsis

Runs 7-Zip so it inherits the console, recording its exit code in
$script:LastSevenZipExit.

### Description

7-Zip writes its password prompt to stdout and only prompts when stdout
is a terminal. Any pipe, redirection or output capture here would hide
the prompt and deadlock, so 7-Zip must inherit this script's console and
its progress output necessarily shares that console. The exit code is
returned through a script variable rather than the pipeline to avoid
capturing (and thereby redirecting) that output.

Original line 354: Prevent PowerShell 7.3+ from turning a non-zero native exit into a

Original line 355: terminating error before we can classify the 7-Zip exit code.

## Resolve-SevenZipOutcome

### Synopsis

Classifies a 7-Zip exit code into a structured outcome.

## Assert-SevenZipOutcome

### Synopsis

Throws a typed, categorised error when 7-Zip did not succeed.

## New-CleanOutputDirectory

### Synopsis

Ensures a destination directory exists, applying -Force overwrite rules.

Original line 436: Resolve to a full path without requiring existence so -WhatIf (which skips

Original line 437: directory creation below) still returns a usable path.

## ConvertTo-Base64Shard

### Synopsis

Writes a binary 7-Zip volume as a single-line base64 text shard.

## ConvertFrom-Base64Shard

### Synopsis

Restores a binary 7-Zip volume from a base64 text shard.

## Resolve-ShardManifest

### Synopsis

Resolves the ordered base64 shard set ('<base>.7z.NNN.txt') from a file
or folder path and returns the shards with their base archive name.

## Invoke-EmbelishOperation

### Synopsis

Packs a folder into an encrypted, split 7-Zip archive, storing the
summary in $script:OperationResult.

Original line 626: Remove each intermediary binary volume as soon as its shard exists.

## Invoke-DeembelishOperation

### Synopsis

Extracts an encrypted, split 7-Zip archive back into a folder, storing
the summary in $script:OperationResult.

Original line 698: No -p: with the password omitted, 7-Zip prompts "Enter password:" on the

Original line 699: console (passing -p with no value would instead use an empty password).

Original line 700: -aoa overwrites existing files (allowed only after the -Force checks above).

## ConvertTo-ResultRecord

### Synopsis

Stamps shared fields onto an operation record and returns it as the
single summary object emitted to the success stream.

## Resolve-ExitCodeFromError

### Synopsis

Maps a terminating error to a documented process exit code.

## Entry point

Original line 775: --- Entry point (the only place that owns 'exit') ---

Original line 779: Keep -Debug non-prompting (5.1 would otherwise set Inquire).

Original line 813: Report non-terminating so the mapped exit code (not the default 1) is used.

## Revision 1.2.0 corrections

EB-01–EB-06 are implemented in the standalone file. `New-CleanOutputDirectory` remains available for its original helper help; the business operations now use `Get-EmbelishDestination` and `New-OutputTransaction` without predeleting final output. `Get-OrderedVolume` validates canonical contiguous nonempty numbering and output counts, and `Resolve-ShardManifest` honors explicit set selection. `Copy-Base64Stream` bounds memory while retaining the original encoding. Native exit 1 is rejected as incomplete mandatory work. `Write-DiagnosticEvent` persists only typed allowlisted fields; the boundary omits raw exception text and emits a stable input/operational/unexpected category.

`Get-EmbelishPathIdentity` contains a narrow Windows adapter for normalized names and FILE_ID_INFO identities. Volume/file IDs are compared against existing ancestors to detect overlap even through alternate drive/UNC/short-name representations; canonical names also stabilize lock filenames. File-identity API failure causes refusal. Directory-link ancestry is separately refused because directory replacement and lock ownership must not be redirected during a run. Unix development uses normalized paths with linked ancestry refused; arbitrary bind-mount identity is not certified.

`Set-OutputAccess` protects staging, copies existing Windows access DACLs as protected snapshots and Unix mode bits, and fails rather than silently publishing with unknown permissions. This changes inheritance and does not clone ownership/audit/alternate-stream metadata; see operations. Its Unix-only reflection selects .NET 6 Unix-mode members absent from .NET Framework, never a Windows fallback. `Copy-OutputTree` streams into existing prepared files to retain their permissions. `Publish-OutputTransaction` rechecks a copied byte inventory and moved previous tree, preserves unrelated entries and retains the entire original on replacement. `Close-OutputTransaction` attempts conditional rollback and reports retained recovery artifacts without masking a primary failure. No claim of atomic multi-directory publication or exclusion of external writers is made.

Two narrowly scoped ShouldProcess analyzer exceptions apply to `New-OutputTransaction` and `Set-OutputAccess`: both are private adapters invoked after the business operation's ShouldProcess gate, and neither may independently skip required staging protection. Preview creates no business staging/lock/output; optional diagnostic I/O remains disclosed.

### Function-named engineering exceptions

The current AST convention counts clauses/loops/catches/Boolean operators, maximum nested control containers, and nonblank token body lines excluding parameter declarations. Here-string native source lines are included conservatively. Parameters remain within six. Focused exceptions retain cohesive validation/publication boundaries rather than obscure their order:

| Function | Exception | Reason |
|---|---|---|
| Write-DiagnosticEvent | Complexity 15 | Explicit fixed allowlist and primitive type checks are privacy controls in a compact body. |
| Get-OrderedVolume | Complexity 13 | Bounds, parse, canonical identity, nonempty data, duplicates and gaps are checked together before mutation. |
| Get-EmbelishPathIdentity | Body 83 lines, including native source | The single-file adapter keeps reviewed signatures, identity/name checks and resource ownership together; individual C# methods are short. |
| Get-EmbelishDestination | Complexity 13 | Platform identity, overlap, root/file and Force preflight checks form one safety decision. |
| Copy-OutputTree | Nesting 4 | A per-file stream try/finally inside inventory/type handling closes both streams even on failure. |
| Publish-OutputTransaction | Complexity 16 | Snapshot/copy/merge, access restoration, prepublication recheck and retained previous-generation ordering must remain reviewable together. |

Native references: [GetFinalPathNameByHandle](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getfinalpathnamebyhandlew), [GetFileInformationByHandleEx](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getfileinformationbyhandleex), [FILE_ID_INFO](https://learn.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-file_id_info), [Set-Acl, PowerShell 5.1](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-acl?view=powershell-5.1).

## Revision 1.2.1 focused corrections

`Read-Base64Stream` uses bounded character buffers, a carry of at most three characters and native Convert.FromBase64String for complete groups. A terminal-padding flag rejects later data across buffers. This restores strict decoding without the whole-file allocation of the original implementation or the silent partial-group discard of FromBase64Transform. `Copy-Base64Stream` retains the encoder and closes its file streams on decode failure. `New-OutputTransaction` tracks whether this invocation created the directory before cleaning up initialization failures; preexisting recovery data is never removed by that catch. `Set-OutputAccess` passes Confirm:false to mandatory Set-Acl under the already approved business operation; errors still terminate. No new analyzer exception, native API or executable dependency is added.

`Get-EmbelishDestination` checks the full filesystem root before trimming separators or resolving native identity. Its complexity is now 13; the existing function-named validation exception covers this additional input check. The root guard is tested locally; the Windows `C:\`/`C:` distinction remains a static platform finding until a real Windows session is available.

## Revision 1.2.2 permission incident and candidate review

The supplied 1.2.1 logs locate three post-compression failures at Set-OutputAccess's Set-Acl call. The supplied external `embelish-fixed.ps1` is reviewed as a candidate, not copied wholesale. Its focused access-control persistence idea is adopted. Raw message/stack/error-ID logging, path-bearing boundary errors and missing-source path disclosure are excluded because they violate the existing redaction contract. Its unrelated file-rule inheritance adjustment is not needed by the current operation paths, and the repository's original encoding/newlines are retained. No supplied log establishes successful Windows execution of the candidate.

`Set-OutputAccess` still builds a new access-only security descriptor, using a private current-user rule or an existing access SDDL/protected snapshot. The writer now calls the documented .NET Framework Directory.SetAccessControl/File.SetAccessControl methods directly, which persist modified descriptor sections. It never sets owner/group/audit data. The Windows PowerShell 5.1 compatibility profile contains both exact string/security overloads. On PowerShell Core, the equivalent FileSystemAclExtensions setter is selected reflectively so that the Core-only type does not become a 5.1 type dependency; provider PSObject wrappers are explicitly unwrapped to native FileSystemInfo/security objects before invocation. A missing setter fails explicitly. Reflection is a runtime compatibility adapter, not a fallback after denied permissions. Both runtime branches propagate native permission errors and have no separate confirmation after outer approval. The existing Unix path is unchanged. A narrow function complexity exception keeps Windows edition/type dispatch and Unix permission compatibility in this single private adapter.

`Write-DiagnosticEvent` adds only an anchored fixed-format HRESULT field; the process boundary derives it from the base exception and keeps the generic error message. This improves evidence without importing the candidate's raw diagnostic fields. Exit-code classification, mandatory protection, staging cleanup, rollback, backup retention, native password I/O and publication ordering remain unchanged.

API references: [Directory.SetAccessControl, .NET Framework](https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.setaccesscontrol?view=netframework-4.8), [File.SetAccessControl, .NET Framework](https://learn.microsoft.com/en-us/dotnet/api/system.io.file.setaccesscontrol?view=netframework-4.8), [FileSystemAclExtensions.SetAccessControl](https://learn.microsoft.com/en-us/dotnet/api/system.io.filesystemaclextensions.setaccesscontrol). Signature checks used installed runtime/profile metadata; no external documentation/network lookup or Windows API execution was performed. See [current verification](verification.md#windows-permission-failure-follow-up-22-october-2026--version-122) for the boundary of local evidence.
