# Calendar Git update helper

[Documentation index](README.md)

Implementation: [`update-calendar.ps1`](../../outcalget/update-calendar.ps1).

## Extracted setup notes

Place this file beside outcalget.ps1 in your Git working tree.
Adjust only this filename if outcalget.ps1 writes a different file.

## Existing behavior and prerequisites

Run explicitly with Windows PowerShell in a Git working tree containing both scripts. Git must be installed and the current branch must have an approved push destination and authentication. The helper sets its location to `$PSScriptRoot` and expects `calendar.ics` there. Configure outcalget first so its saved output path matches that filename, or deliberately adjust `$calendarFile` to the approved output. This helper does not create an export configuration.

```powershell
powershell.exe -NoLogo -NoProfile -File "C:\ApprovedTools\outcalget\update-calendar.ps1"
```

The helper removes an existing calendar file, invokes `outcalget.ps1` with `-STA`, checks the exit code and output, stages only the calendar, commits it if changed, and runs `git push`. Errors exit with code 1. A push failure leaves the commit local.

## Review findings

This documentation extraction preserves the helper as implemented. Its delete-before-export sequence can discard the previous calendar if export fails, and usually defeats the exporter's existing-output protection. It also intentionally publishes full calendar data through Git: use only an approved destination. `git push` follows the repository's existing branch/remote configuration. These effects differ from running the exporter alone; this helper has no `-WhatIf` contract. Do not execute it during documentation or sandbox checks.

Changing deletion, publication, prompting or configuration behavior is a separate implementation change requiring focused review and validation.

The helper's two existing `Write-Host` calls produce `PSAvoidUsingWriteHost` analyzer warnings. They remain unchanged in this extraction; the analyzer regression comparison confirms that no new findings were introduced. This does not claim a clean analyzer result for the helper.
