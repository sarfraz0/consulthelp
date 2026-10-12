# Embelish and Sendshards instructions

Read the root `AGENTS.md` and `docs/Powershell guidelines.md`. For `embelish.ps1`, read `docs/embelish/README.md`; for `sendshards.ps1`, read `docs/sendshards/README.md`. Both tools share this implementation directory but have separate documentation and behavior contracts.

## Embelish contracts

- Preserve the 7-Zip-native encrypted/header-encrypted split archive and single-line base64 shard format, established parameter validation, default compression level 9 and 10 MiB binary volume setting.
- Preserve the interactive password flow. 7-Zip inherits the console; do not redirect/capture its output, add passwords to arguments/configuration/environment/logs, or introduce unattended credentials as part of unrelated work.
- Preserve both modes, direct native exit-code classification, `-Force` rules, `ShouldProcess`/`-WhatIf`, result objects, diagnostic behavior and staging cleanup. Do not add silent overwrite or retry behavior.
- Read the documented limitations before claiming authenticity, sanitized diagnostics, preview with no diagnostic I/O, contiguous shard validation, multi-set selection or previous-output preservation. The current producer emits no checksum manifest.
- Do not turn the documentation extraction into a storage, crypto, password-management or security-inspection change. Base64 is data encoding, not a way to bypass corporate controls.

## Sendshards contracts

- Preserve distinct `-Plan` and `-Push` modes. Planning performs local checks and publishes a new plan without network, staging, commit or push; repeating Push resumes the same transfer.
- Preserve frozen SHA-256/Git blob/source/index/commit checks, branch/base/remote identity, exact one-file commits, manifest-after-shards ordering, legacy-set warnings, source retention and checkpoint reconciliation with actual Git/remote state.
- Preserve deterministic rate limits, the original persisted 60-minute budget, per-command deadlines, bounded captured output, exclusive common-directory lock and process-tree cleanup limited to owned processes. Never restart the budget silently or blindly repeat an uncertain push.
- Keep plans/state/logs outside worktrees and Git metadata; preserve link-ancestry and schema checks, byte-changing filter rejection, and refusal of foreign staged changes, branch divergence and changed inputs.
- Preserve native hooks, signing, credential helpers, TLS/proxy checks and repository filters. Do not reset, stash, rewrite history, force-push, skip hooks, bypass inspection or add randomized pacing.
- Preserve no success-stream result, safe diagnostics, exit codes and read-only `WhatIf`/declined confirmation semantics. Actual transfer requires Windows, approved Git and pre-provisioned credentials; Linux analysis does not prove it works on the target.

## Documentation and checks

- Both scripts remain focused on implementation. Preserve `#requires`, `.EXTERNALHELP`, executable tokens/string literals, original encodings and CLI contracts. Public and Embelish function help live in `docs/embelish/embelish-help.xml`; Sendshards help lives in `docs/sendshards/sendshards-help.xml`.
- Keep Markdown and XML synchronized and verify `Get-Help -Full` and `-?` from an unrelated directory. Embelish's 15 existing helper help blocks must remain accessible through native help.
- Compare executable token kind/text when extracting documentation. Load selected AST function definitions with their original source filename to check helper help; never dot-source the complete scripts because their entry points perform work and exit.
- Use isolated synthetic files/repositories if behavior testing is needed. Do not invoke real archive operations, password prompts, Sendshards Push or calendar publishing merely to validate documentation. User authorization to edit code does not authorize publication of data.
- Keep historical evidence separate from current parser/analyzer/help results; report existing analyzer findings rather than suppressing them to get a clean result. Record Windows, console, scheduler, endpoint and live-client acceptance as unverified when unavailable.
