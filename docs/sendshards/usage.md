# sendshards usage

[Documentation index](README.md)

## Parameters

### -Plan

Generate a new plan. Existing plan or checkpoint files are never overwritten.

### -Push

Execute or resume an existing plan. Repeating this command is the resume action.

### -SourcePath

With Plan: a local folder inside the worktree, containing *.7z.NNN.txt shards.
Top-level files only; multiple archive sets are supported. Other files ignored.

### -PlanPath

Required in either mode. JSON document outside the repository. Relative paths
use the invocation's PowerShell filesystem location, not the script directory.

### -Remote

With Plan: configured remote name, default origin. Exactly one push URL required.
The configured push URL is used for both push and remote verification.

### -WaitSeconds

With Plan: fixed minimum wait between pushes, default 30, range 5 through 600.
This is a rate limit, not a means of avoiding security inspection.

### -GitTimeoutSeconds

With Plan: maximum duration of each Git command, default 60, range 10 through 300,
further limited by the remaining execution budget.

### -GitPath

Optional approved installed git.exe path. Otherwise resolved from PATH.

## Examples

### Example 1

```powershell
.\sendshards.ps1 -Plan -SourcePath 'C:\repo\shards' -PlanPath 'C:\TransferPlans\batch.json'
```



### Example 2

```powershell
.\sendshards.ps1 -Push -PlanPath 'C:\TransferPlans\batch.json'
```

Repeat the same command after an interruption; no separate Resume flag is needed.

### Example 3

```powershell
.\sendshards.ps1 -Push -PlanPath 'C:\TransferPlans\batch.json' -WhatIf
```



### Example 4

```powershell
.\sendshards.ps1 -Push -PlanPath 'C:\TransferPlans\batch.json' -Debug
```



## Inputs

None. Source files and the persisted plan are input data.

## Outputs

None on the success stream. Plan produces a JSON document. Push maintains a
checkpoint and sanitized JSONL operational logs. Verbose/Debug use their streams.
