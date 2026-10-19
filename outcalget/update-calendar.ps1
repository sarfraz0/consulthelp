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
    param([string] $Text)
    $result = ConvertFrom-Json -InputObject $Text -ErrorAction Stop
    if ($result -isnot [pscustomobject]) {
        throw 'outcalget.ps1 did not return an export result.'
    }
    foreach ($key in @('resultSchemaVersion', 'operation', 'exitCode', 'published', 'outputPath')) {
        if ($null -eq $result.PSObject.Properties[$key]) {
            throw 'outcalget.ps1 returned an incomplete export result.'
        }
    }
    if (($result.resultSchemaVersion -isnot [int] -and $result.resultSchemaVersion -isnot [long]) -or $result.resultSchemaVersion -ne 1 -or
        $result.operation -isnot [string] -or $result.operation -cne 'export' -or
        ($result.exitCode -isnot [int] -and $result.exitCode -isnot [long]) -or $result.exitCode -ne 0 -or
        $result.published -isnot [bool] -or -not $result.published -or
        $result.outputPath -isnot [string] -or [string]::IsNullOrWhiteSpace($result.outputPath)) {
        throw 'outcalget.ps1 did not confirm a successfully published calendar.'
    }
    if (-not [System.IO.Path]::IsPathRooted($result.outputPath) -or
        -not (Test-Path -LiteralPath $result.outputPath -PathType Leaf)) {
        throw 'The published calendar path is not an existing absolute file.'
    }
    return (Get-Item -LiteralPath $result.outputPath).FullName
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
        return
    }
    $output = & $hostPath @arguments
    $exportCode = $LASTEXITCODE
    if ($exportCode -ne 0) {
        throw "outcalget.ps1 failed (exit code $exportCode). No commit or push was attempted."
    }
    $calendar = Get-PublishedCalendarPath -Text ($output -join "`n")
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
