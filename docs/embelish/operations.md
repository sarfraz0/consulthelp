# embelish operations

[Documentation index](README.md)

## Extracted operational notes

Exit codes:
  0  Operation completed successfully.
  1  Unexpected or unclassified failure.
  2  Invalid input or configuration (bad arguments, missing source,
     existing output without -Force, 7-Zip command-line error).
  3  Operational, dependency or permission failure (7-Zip not found, 7-Zip
     fatal error, wrong password, I/O or access error).
  4  Reserved for a controlled timeout (not used by this script).

Scheduled / Task Scheduler note:
  Because 7-Zip prompts for the password on the console, this tool needs an
  interactive session. A classic Task Scheduler "Program/script" action is:
    Program/script: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
    Arguments:      -NoLogo -NoProfile -File "C:\ApprovedTools\embelish.ps1" -Mode Embelish -SourcePath "D:\data" -DestinationPath "E:\vault"
    Start in:       C:\ApprovedTools
  Do not use silent scheduling for this interactive workflow. Append
  -Debug for extra diagnostics. Paths above are illustrative.

Requirements:
  Windows PowerShell 5.1 or PowerShell 7+, and a 7-Zip executable on the
  target machine (set -SevenZipPath if it is not on PATH).

## Existing implementation limits and documentation corrections

The archive is encrypted by 7-Zip; this script does not implement a separate authenticity check or establish authenticated encryption. The former description's authenticated-encryption wording has been corrected to archive encryption in Markdown and native help. Base64 expands binary volume bytes by approximately one third; `VolumeSizeMegabytes` controls the binary 7-Zip volume, not the encoded `.txt` file's size.

`-WhatIf` prevents the archive operation, but diagnostic initialization and retention occur before that gate. It can create/write/prune logs, and it still resolves the 7-Zip dependency. The historical claim of no mutation refers to archive behavior; it must not be used to promise no diagnostic I/O.

`-Force` can remove old shards before a new archive succeeds, and failed output conversion can leave partial shards. Staging cleanup is attempted in `finally` with suppressed cleanup errors. This extraction preserves those behaviors; it does not claim atomic replacement or guaranteed cleanup.

`Resolve-ShardManifest` orders the discovered shards and checks that numbering starts at 001; it does not verify every gap or publish a checksum manifest. Multiple sets in the same directory are refused even if `SourcePath` names one shard, because the current implementation still examines the entire containing directory. Keep sets in separate directories. Sendshards applies its own stricter contiguous-numbering and byte checks and accepts these manifest-free sets with a legacy warning.

The password prompt belongs to the native 7-Zip process. PowerShell's `-NonInteractive` is not a general control over native prompts. Use an interactive console and do not promise unattended Task Scheduler operation without separately designed credential handling.

Passwords are not passed to the logger, but logs can include `sourceLeaf` and raw exception messages. Warning/error output can also expose paths. Review logs and outputs before sharing; the current logger is not claimed to meet every redaction default in the shared guidelines. Sensitive real archive data is not needed to verify this extraction.
