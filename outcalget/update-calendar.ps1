#requires -Version 5.1

# .EXTERNALHELP ../docs/outcalget/update-calendar-help.xml
[CmdletBinding(PositionalBinding = $false, SupportsShouldProcess = $true, ConfirmImpact = 'Low', DefaultParameterSetName = 'Profile')]
param(
    [string] $ExporterPath = (Join-Path -Path $PSScriptRoot -ChildPath 'outcalget.ps1'),
    [Parameter(ParameterSetName = 'Config')]
    [string] $Config,
    [Parameter(ParameterSetName = 'Profile')]
    [Alias('Profile')]
    [string] $ExportProfile
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

function Get-PublishedCalendarPath {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Result)
    foreach ($key in @('resultSchemaVersion', 'operation', 'exitCode', 'published', 'outputPath')) {
        if ($null -eq $result.PSObject.Properties[$key]) {
            throw 'outcalget.ps1 returned an incomplete export result. No commit or push was attempted.'
        }
    }
    if (($result.resultSchemaVersion -isnot [int] -and $result.resultSchemaVersion -isnot [long]) -or $result.resultSchemaVersion -ne 1 -or
        $result.operation -isnot [string] -or $result.operation -cne 'export' -or
        ($result.exitCode -isnot [int] -and $result.exitCode -isnot [long]) -or $result.exitCode -ne 0 -or
        $result.published -isnot [bool] -or -not $result.published -or
        $result.outputPath -isnot [string] -or [string]::IsNullOrWhiteSpace($result.outputPath)) {
        throw 'outcalget.ps1 did not confirm a successfully published calendar. No commit or push was attempted.'
    }
    if (-not [System.IO.Path]::IsPathRooted($result.outputPath) -or
        -not (Test-Path -LiteralPath $result.outputPath -PathType Leaf)) {
        throw 'The published calendar path is not an existing absolute file. No commit or push was attempted.'
    }
    return (Get-Item -LiteralPath $result.outputPath).FullName
}

function Get-CalendarExportMessage {
    [CmdletBinding()]
    [OutputType([string])]
    param([object] $Detail)
    if ($Detail -isnot [pscustomobject]) {
        return
    }
    $code = $Detail.PSObject.Properties['code']
    $message = $Detail.PSObject.Properties['message']
    if ($null -ne $code -and $code.Value -is [string] -and $null -ne $message -and $message.Value -is [string]) {
        return ('[{0}]: {1}' -f $code.Value, $message.Value)
    }
}

function Write-CalendarExportNotice {
    [CmdletBinding()]
    param([object] $Result)
    $record = $Result.PSObject.Properties['runLogPath']
    if ($null -ne $record -and $record.Value -is [string] -and -not [string]::IsNullOrWhiteSpace($record.Value)) {
        Write-Information -InformationAction Continue -MessageData ('Outlook export run record: ' + $record.Value)
    }
    $warnings = $Result.PSObject.Properties['warnings']
    if ($null -eq $warnings -or $warnings.Value -isnot [array]) {
        return
    }
    foreach ($warning in $warnings.Value) {
        $message = Get-CalendarExportMessage -Detail $warning
        if (-not [string]::IsNullOrEmpty($message)) {
            Write-Warning -Message $message
        }
    }
}

function Read-CalendarExportResult {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $Text, [int] $ExitCode)
    try {
        $result = ConvertFrom-Json -InputObject $Text -ErrorAction Stop
    } catch {
        throw "outcalget.ps1 returned no valid JSON export result (exit code $ExitCode). No commit or push was attempted."
    }
    if ($result -isnot [pscustomobject]) {
        throw "outcalget.ps1 did not return an export result (exit code $ExitCode). No commit or push was attempted."
    }
    Write-CalendarExportNotice -Result $result
    if ($ExitCode -ne 0) {
        $failure = "outcalget.ps1 failed (exit code $ExitCode)."
        $errorProperty = $result.PSObject.Properties['error']
        if ($null -ne $errorProperty) {
            $message = Get-CalendarExportMessage -Detail $errorProperty.Value
            if (-not [string]::IsNullOrEmpty($message)) {
                $failure += ' ' + $message
            }
        }
        $published = $result.PSObject.Properties['published']
        if ($null -ne $published -and $published.Value -is [bool] -and $published.Value) {
            $failure += ' Output was published before the later failure; inspect the run record before retrying.'
        }
        throw ($failure + ' No commit or push was attempted.')
    }
    return $result
}

function Get-CalendarRepository {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $GitPath, [string] $CalendarPath)
    $directory = [System.IO.Path]::GetDirectoryName($CalendarPath)
    $rootOutput = & $GitPath -C $directory rev-parse --show-cdup
    if ($LASTEXITCODE -ne 0) {
        throw 'The published calendar is not inside a Git working tree. No commit or push was attempted.'
    }
    $root = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($directory, ($rootOutput -join '')))
    $prefix = $root.TrimEnd([char[]]'\/') + [System.IO.Path]::DirectorySeparatorChar
    if (-not $CalendarPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'The published calendar is outside the resolved Git working tree.'
    }
    return [pscustomobject]@{
        Root = $root
        RelativePath = $CalendarPath.Substring($prefix.Length).Replace('\', '/')
    }
}

try {
    $exporter = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExporterPath)
    if (-not (Test-Path -LiteralPath $exporter -PathType Leaf)) {
        throw 'outcalget.ps1 was not found. Place it beside this script or supply -ExporterPath.'
    }
    $gitPath = (Get-Command -Name git -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $hostPath = (Get-Command -Name powershell.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-OutputFormat', 'XML', '-File', $exporter, '-Json')
    if ($PSBoundParameters.ContainsKey('Config')) {
        $configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
        $arguments += @('-Config', $configPath)
    }
    if ($PSBoundParameters.ContainsKey('ExportProfile')) {
        $arguments += @('-Profile', $ExportProfile)
    }
    if (-not $PSCmdlet.ShouldProcess('The configured calendar and its Git repository', 'Export, commit the calendar if changed, and push')) {
        Write-Information -InformationAction Continue -MessageData 'Calendar update skipped. No export, commit or push was attempted.'
        return
    }
    Write-Verbose -Message ('Exporter: ' + $exporter)
    Write-Information -InformationAction Continue -MessageData 'Starting Outlook calendar export. Waiting for outcalget to finish...'
    $output = & $hostPath @arguments
    $exportCode = $LASTEXITCODE
    $result = Read-CalendarExportResult -Text ($output -join "`n") -ExitCode $exportCode
    $calendar = Get-PublishedCalendarPath -Result $result
    Write-Information -InformationAction Continue -MessageData ('Outlook calendar export completed: ' + $calendar)
    $repository = Get-CalendarRepository -GitPath $gitPath -CalendarPath $calendar
    $gitArguments = @('--literal-pathspecs', '-C', $repository.Root)
    & $gitPath @gitArguments add -- $repository.RelativePath
    if ($LASTEXITCODE -ne 0) {
        throw 'git add failed.'
    }
    & $gitPath @gitArguments diff --cached --quiet -- $repository.RelativePath
    $diffCode = $LASTEXITCODE
    if ($diffCode -eq 0) {
        Write-Output -InputObject 'Calendar unchanged. Nothing to commit or push.'
    } elseif ($diffCode -eq 1) {
        & $gitPath @gitArguments commit --only -m 'Update calendar' -- $repository.RelativePath
        if ($LASTEXITCODE -ne 0) {
            throw 'git commit failed. Nothing was pushed.'
        }
        & $gitPath -C $repository.Root push
        if ($LASTEXITCODE -ne 0) {
            throw 'git push failed. The commit remains local.'
        }
        Write-Output -InputObject 'Calendar updated, committed and pushed.'
    } else {
        throw "git diff failed (exit code $diffCode)."
    }
} catch {
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
    exit 1
}
