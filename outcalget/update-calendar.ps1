# Place this file beside outcalget.ps1 in your Git working tree.
# Adjust only this filename if outcalget.ps1 writes a different file.
$calendarFile = 'calendar.ics'
$ErrorActionPreference = 'Stop'

try {
    Set-Location -LiteralPath $PSScriptRoot

    if (-not (Test-Path -LiteralPath '.\outcalget.ps1' -PathType Leaf)) {
        throw 'outcalget.ps1 was not found beside this script.'
    }

    git rev-parse --is-inside-work-tree | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'This folder is not inside a Git working tree.' }

    if (Test-Path -LiteralPath $calendarFile -PathType Leaf) {
        Remove-Item -LiteralPath $calendarFile -Force
    }

    & powershell.exe -NoLogo -NoProfile -STA -File (Join-Path $PSScriptRoot 'outcalget.ps1')
    if ($LASTEXITCODE -ne 0) { throw "outcalget.ps1 failed (exit code $LASTEXITCODE)." }

    if (-not (Test-Path -LiteralPath $calendarFile -PathType Leaf)) {
        throw "outcalget.ps1 did not create $calendarFile. No commit or push was attempted."
    }

    git add -- $calendarFile
    if ($LASTEXITCODE -ne 0) { throw 'git add failed.' }

    git diff --cached --quiet -- $calendarFile
    $diffCode = $LASTEXITCODE
    if ($diffCode -eq 0) {
        Write-Host 'Calendar unchanged. Nothing to commit or push.'
    }
    elseif ($diffCode -eq 1) {
        git commit --only -m 'Update calendar' -- $calendarFile
        if ($LASTEXITCODE -ne 0) { throw 'git commit failed. Nothing was pushed.' }

        git push
        if ($LASTEXITCODE -ne 0) { throw 'git push failed. The commit remains local.' }

        Write-Host 'Calendar updated, committed and pushed.'
    }
    else {
        throw "git diff failed (exit code $diffCode)."
    }
}
catch {
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
    exit 1
}
