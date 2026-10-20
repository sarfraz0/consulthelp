# embelish

Protects (embelishes) or restores (deembelishes) a folder as encrypted,
base64-encoded text shards backed by a split 7-Zip archive.

Implementation: [`embelish.ps1`](../../embelish/embelish.ps1), version 1.2.0.

**Production readiness: target acceptance pending.** Version 1.2.0 implements EB-01–EB-06 from the [20 October audit](audit-2026-10-20.md). Current synthetic verification and recovery instructions are recorded below; Windows PowerShell 5.1, Windows filesystem/ACL behavior and managed-endpoint acceptance remain unverified.

## Documentation

- [Parameters, examples and results](usage.md)
- [Operation, scheduling and existing limits](operations.md)
- [Function documentation and implementation rationale](implementation-notes.md)
- [Historical verification and current extraction](verification.md)
- [20 October production audit and release gates](audit-2026-10-20.md)
- [Native PowerShell help](embelish-help.xml)

## Behavior and prerequisites

Embelish packs a source folder into an AES-256 encrypted, header-encrypted,
multi-volume 7-Zip archive, then stores each volume as a single-line base64
text shard ('<name>.7z.001.txt', '.002.txt', ...). Deembelish reverses the
operation: it base64-decodes the shards back into 7-Zip volumes and extracts
them into a folder. Compression, archive encryption, filename hiding
and volume splitting are delegated to the 7-Zip command-line tool; a
private staging tree beside the destination holds intermediate volumes and
restored files. Publication uses same-parent directory renames and retains
the previous destination; see [recovery and storage requirements](operations.md).

The encryption password is never accepted on the command line, stored in a
file, placed in an environment variable or written to any log. Instead, the
7-Zip tool prompts for it directly on the console. 7-Zip owns password entry; prompt count depends on the installed build.
The audited Linux 25.01 creation/extraction invocations each prompted once;
Windows console behavior remains unverified. Because of this the
script is interactive by design and is not intended for silent, unattended
scheduling unless a separate provisioned credential mechanism is added.

Diagnostic events are written as JSON Lines to a per-run file under
%LOCALAPPDATA%\Embelish\logs (or -LogDirectory). The diagnostic file records
allowlisted stage/count/exit/error-category fields. Paths, source names, raw
exceptions, payloads and passwords are omitted. Native 7-Zip console output
can contain filenames; it is not a sanitized log.

## Deployment and help

The script can live in `scripts/`, a repository root, or an independent tools directory. Relative source, destination, explicit 7-Zip and log paths resolve from the caller’s current PowerShell filesystem location, independent of the script directory. Version 1.1.1 corrects relative log-path resolution when PowerShell’s location differs from the process directory.

Deploy `embelish.ps1` alone with its declared PowerShell/7-Zip dependencies and an interactive console. No XML, Markdown, agent instructions or companion script is required for archive/restore execution. Optional XML supplies detailed native help; without it, `Get-Help` and `-?` fall back to generated syntax. To retain detailed help, also copy matching `embelish/` and `docs/` directories together so the native `.EXTERNALHELP` directive resolves `../docs/embelish/embelish-help.xml`. Keep Markdown and native XML help synchronized when documentation changes.

## Related tool

[Sendshards](../sendshards/README.md) can plan transfer of the produced shards. This Embelish revision emits no checksum manifest; Sendshards treats those sets as legacy and warns.

Version 1.2.0 requires FullLanguage and installed 7-Zip. On Windows it uses built-in native file-identity APIs for alias-safe overlap checks and standard ACL cmdlets. Filesystems must support these checks and same-parent directory renames. Unix development execution requires PowerShell 7.2/.NET 6 or newer for private staging modes. No extra modules, XML, documentation or companion implementation files are required.
