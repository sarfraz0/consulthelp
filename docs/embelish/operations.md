# embelish operations

[Documentation index](README.md)

## Production audit status

Version **1.2.0** implements the six findings in the [20 October audit](audit-2026-10-20.md). [Current verification](verification.md) separates passing synthetic checks from Windows/endpoint acceptance, which remains pending.

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
  Windows PowerShell 5.1 in FullLanguage (Unix development: PowerShell 7.2/.NET 6+), and a 7-Zip executable on the
  target machine (set -SevenZipPath if it is not on PATH).

## Existing implementation limits and documentation corrections

The archive is encrypted by 7-Zip; this script does not implement a separate authenticity check or establish authenticated encryption. The former description's authenticated-encryption wording has been corrected to archive encryption in Markdown and native help. Base64 expands binary volume bytes by approximately one third; `VolumeSizeMegabytes` controls the binary 7-Zip volume, not the encoded `.txt` file's size.

`-WhatIf` prevents the archive operation, but diagnostic initialization and retention occur before that gate. It can create/write/prune logs, and it still resolves the 7-Zip dependency. The historical claim of no mutation refers to archive behavior; it must not be used to promise no diagnostic I/O.

## Staging, publication and recovery

Allow replacement only after successful staging. Embelish replaces the selected archive base and removes its obsolete trailing shards; Deembelish merges restored entries into the existing tree. Unrelated files are retained. The entire previous destination is retained in a sibling backup, reported as BackupPath. Force does not bypass locks, validation or ShouldProcess.

Both operations reject source/output overlap before business mutations. Windows checks canonical paths and native volume/file identities, including ancestors, to catch drive/short-name aliases. Use actual directory paths: symbolic-link/junction ancestry is refused at this publication boundary. Copied/extracted output trees refuse reparse points and hard-linked files because following them would invalidate the snapshot and could publish outside the owned tree. Hydrate cloud placeholders or choose an ordinary output directory when this check refuses an entry; this is a tree-publication requirement, not a general Windows filesystem prohibition.

After ShouldProcess approves, the tool holds an exclusive sibling `.NAME.embelish.lock` file and prepares `.NAME.embelish.transaction` on the destination filesystem. Native archive creation/extraction and base64 conversion touch only that staging tree. Native exit **1 is now an operational failure**, since skipped source entries cannot be reported as complete mandatory work. Volume sets must contain 1–10,000 nonempty, canonical, contiguous volumes starting at 001. An explicit shard file selects its exact archive set; a directory containing several sets remains ambiguous. Base64 conversion streams bytes with bounded buffers and preserves the existing single-line format.

Before publication, an existing destination is copied to the candidate, checked against its byte inventory, and merged with incoming entries. Immediately before replacing it, it is renamed to `previous` inside the transaction and rechecked. A no-Force late nonempty destination is refused. The prepared tree is moved into place without replacing any newly created destination. There is a short directory-rename gap; this is **recoverable publication, not a filesystem-wide atomic transaction**. Other applications are not locked out: stop writers/readers that rename or modify these trees during a run. A late writer remains in the retained previous tree or its own newly created destination; inspect both after a publication conflict.

On successful replacement of any existing destination, `BackupPath` names `.NAME.embelish.transaction.backup-<id>` beside it. Its `previous` child is the entire old tree. Backups are never automatically pruned. For a new destination, BackupPath is null. Successful backups contain no extra native/incoming staging copies. Plan storage for the old tree, its prepared merged copy, encoded shards, native volumes and extraction output; disk exhaustion fails the operation and leaves existing output untouched or retained for recovery.

If publication has begun and fails, rollback is attempted only when the destination is absent. Recovery artifacts are retained and subsequent runs refuse the fixed transaction directory. With all related processes stopped, inspect the requested destination and `.NAME.embelish.transaction/{previous,ready}`. Copy valuable late writes aside first. If the destination is absent, move `previous` back to the destination; if it already contains valuable output, preserve both trees and choose the intended generation manually. Remove or rename the fixed transaction directory only after recovery is verified. Do not delete `previous` to make a retry start. A leftover lock filename alone is normal: the exclusive open handle, rather than filename existence, establishes ownership.

Windows staging starts with access for the current user only. Existing directory/file DACLs are copied as protected permission snapshots before publication, preventing new parent inheritance from broadening access; new entries inherit their prepared parent. Protected snapshots deliberately no longer inherit future ancestor ACL changes automatically. A new destination is private to the invoking user. Unix staging uses mode 0700 and existing mode bits are copied. Ownership, audit ACLs, alternate streams and all filesystem metadata are not cloned; the retained original tree remains the authoritative previous generation. Target Windows ACL/UNC/short-name and interruption tests are still required.

The password prompt belongs to the native 7-Zip process. PowerShell's `-NonInteractive` is not a general control over native prompts. Use an interactive console and do not promise unattended Task Scheduler operation without separately designed credential handling.

The JSONL logger now allowlists typed operational fields and stable error categories/line numbers. Source paths/names, arbitrary fields and raw exception messages are omitted, including logging-failure warnings. Summary results deliberately still contain source, destination, executable, log and backup paths. Native 7-Zip console diagnostics can contain filenames. Review console captures and result objects before sharing.

Contiguous names do not prove that all final volumes arrived, nor that every intended source entry was captured. The producer still emits no checksum/source-inventory manifest. Successful native extraction plus an expected file/hash inventory is the recipient acceptance check. Archives must come from a trusted sender; this tool is not a sandbox or separate authenticity verifier. Windows console, endpoint, resource-exhaustion and abrupt-termination acceptance remain pending.
