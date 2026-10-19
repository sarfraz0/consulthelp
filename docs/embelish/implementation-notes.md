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
