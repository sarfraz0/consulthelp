# embelish

Protects (embelishes) or restores (deembelishes) a folder as encrypted,
base64-encoded text shards backed by a split 7-Zip archive.

Implementation: [`embelish.ps1`](../../embelish/embelish.ps1), version 1.1.0.

## Documentation

- [Parameters, examples and results](usage.md)
- [Operation, scheduling and existing limits](operations.md)
- [Function documentation and implementation rationale](implementation-notes.md)
- [Historical verification and current extraction](verification.md)
- [Native PowerShell help](embelish-help.xml)

## Behavior and prerequisites

Embelish packs a source folder into an AES-256 encrypted, header-encrypted,
multi-volume 7-Zip archive, then stores each volume as a single-line base64
text shard ('<name>.7z.001.txt', '.002.txt', ...). Deembelish reverses the
operation: it base64-decodes the shards back into 7-Zip volumes and extracts
them into a folder. Compression, archive encryption, filename hiding
and volume splitting are delegated to the 7-Zip command-line tool; a
short-lived staging directory under the system temp folder holds the
intermediate volumes and is removed when the operation finishes.

The encryption password is never accepted on the command line, stored in a
file, placed in an environment variable or written to any log. Instead, the
7-Zip tool prompts for it directly on the console. Embelish asks for the
password twice (enter and verify); deembelish asks once. Because of this the
script is interactive by design and is not intended for silent, unattended
scheduling unless a separate provisioned credential mechanism is added.

Diagnostic events are written as JSON Lines to a per-run file under
%LOCALAPPDATA%\Embelish\logs (or -LogDirectory). The diagnostic file records
operational facts only (operation, sizes, counts, 7-Zip exit code, timing);
it never contains the password.

## Deployment and help

Deploy `embelish.ps1` alone with its declared PowerShell/7-Zip dependencies and an interactive console. No XML, Markdown, agent instructions or companion script is required for archive/restore execution. Optional XML supplies detailed native help; without it, `Get-Help` and `-?` fall back to generated syntax. To retain detailed help, also copy matching `embelish/` and `docs/` directories together so the native `.EXTERNALHELP` directive resolves `../docs/embelish/embelish-help.xml`. Keep Markdown and native XML help synchronized when documentation changes.

## Related tool

[Sendshards](../sendshards/README.md) can plan transfer of the produced shards. This Embelish revision emits no checksum manifest; Sendshards treats those sets as legacy and warns.
