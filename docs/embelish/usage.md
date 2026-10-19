# embelish usage

[Documentation index](README.md)

## Parameters

### -Mode

Operation to perform: 'Embelish' to protect a folder, 'Deembelish' to
restore one.

### -SourcePath

Embelish: the folder to protect.
Deembelish: the folder that holds the base64 shards, or any one shard file
(a '*.7z.NNN.txt') from the set.

### -DestinationPath

Embelish: the output folder that will receive the base64 text shards.
Deembelish: the output folder that will receive the restored files.

### -VolumeSizeMegabytes

Size of each shard in megabytes for Embelish. Default 10. Ignored by
Deembelish (7-Zip reassembles whatever volume sizes it finds).

### -CompressionLevel

7-Zip compression level 0-9 for Embelish (0 = store, 9 = ultra/maximum).
Default 9 (highest). Ignored by Deembelish.

### -SevenZipPath

Full path to the 7-Zip executable. When omitted the script searches PATH for
7z, 7za, 7zz and 7zr, then the standard Windows install locations.

### -ArchiveBaseName

Base file name used for the Embelish shards. The produced shards are
'<ArchiveBaseName>.7z.001.txt', '.002.txt', and so on. Default 'embelish'.
Allowed characters: letters, digits, dot, dash and underscore.

### -LogDirectory

Directory for the JSON Lines diagnostic file. Defaults to
%LOCALAPPDATA%\Embelish\logs. Relative values resolve from the caller’s current PowerShell filesystem location.

### -Force

Overwrite existing output. For Embelish this removes pre-existing shards
that share the archive base name. For Deembelish this allows extracting into
a non-empty destination folder.

## Examples

### Example 1

```powershell
.\embelish.ps1 -Mode Embelish -SourcePath C:\data\project -DestinationPath C:\vault\project
```

Prompts for a password and writes encrypted base64 shards
(embelish.7z.001.txt, ...) into C:\vault\project.

### Example 2

```powershell
.\embelish.ps1 -Mode Deembelish -SourcePath C:\vault\project -DestinationPath C:\restore\project
```

Reads the base64 shards in C:\vault\project, prompts for the password and
restores the folder into C:\restore\project.

### Example 3

```powershell
.\embelish.ps1 -Mode Embelish -SourcePath C:\data\project -DestinationPath C:\vault\project -WhatIf
```

Shows what would happen without creating any archive.

## Inputs

None. This script does not accept pipeline input.

## Outputs

[pscustomobject] A single summary record describing the completed (or
previewed) operation.
