# sendshards

Plans and resumes a paced transfer of archive shards to an approved Git remote.

Implementation: [`sendshards.ps1`](../../embelish/sendshards.ps1), version 1.0.0.

## Documentation

- [Parameters, examples and results](usage.md)
- [Operation, scheduling and existing limits](operations.md)
- [Function documentation and implementation rationale](implementation-notes.md)
- [Historical verification and current extraction](verification.md)
- [Native PowerShell help](sendshards-help.xml)

## Behavior and prerequisites

Revision 1.0.0. One self-contained Windows PowerShell 5.1 script. Requires
installed Git for Windows 2.30 or newer, FullLanguage, and pre-provisioned credentials.
Run as the current user. No installation, elevation, policy changes are performed. Git hooks, signing,
credential helpers, TLS checks, and repository filters remain active.

Plan reads local repository information and writes a readable UTF-8 JSON
document. It performs no network operation, staging, commit, or push. The
document records SHA-256 hashes, exact Git blob IDs, the branch, the configured
push destination's fingerprint, and a fixed wait for each pending file.
Pacing is transparent and deterministic.

Push revalidates that document and processes one pending file at a time:
verify -> add -> verify index -> commit -> verify commit -> push -> verify
remote branch -> checkpoint. All shard sets must have contiguous numbering.
Each matching checksum manifest is included after that set's shards. Legacy
sets are accepted with a warning. This does not decrypt or test the archive.
Already-identical files in the base commit need no new commit.

Push starts a 60-minute budget, persisted at <PlanPath>.state.json. The same
Push command resumes automatically. Verified Git commits and the actual
remote branch, rather than an optimistic counter, establish progress. One
local commit may be awaiting push. Failed/uncertain pushes are not blindly
retried. Last attempted/verified push times preserve pacing on resume.
A timed-out plan remains inspectable but cannot restart its original budget.
A completed checkpoint may be reverified later without sending more data.
A new Plan is an explicit new transfer, with a new budget.
Its base still must match the remote: resolve any pending older local commit
through the approved Git recovery procedure before starting a replacement plan.

Completion within 60 minutes cannot be guaranteed: counts, bandwidth,
authentication, hooks, server limits, and outages can make it impossible.
Plan rejects waits consuming more than 50 minutes, reserving at least ten
minutes for work. Push bounds Git calls and stops initiating work at its
deadline. Blocking filesystem I/O and process cleanup can exceed the budget.
Keep the Windows clock accurate. Fixed waiting starts after the previous
attempt/verified completion; operation time adds to the planned waits.

The current branch must already exist remotely. At the first Push its remote
tip must equal the plan's base commit, preventing upload of unrelated local
commits. Divergence, branch changes, changed source bytes, and foreign staged
changes cause a stop. Do not edit or use Git in this worktree during transfer.
The exclusive sendshards.lock in Git's common directory serializes this tool
across linked worktrees; it cannot lock out other Git applications.

Plan/state/logs must be outside the worktree and Git metadata. Plans contain
private paths and file hashes and must be reviewed before sharing. Logs omit
paths, remote URLs, payloads, credentials, raw Git output, and exception text.
No success-stream result is produced. Files/shards are never removed. Failed
commits can leave the next file staged; successful commits remain in Git.
A verified, matching staged next file is accepted on resume. Rewriting history,
resetting, stashing, force-pushing, and silently skipping hooks are not used.
Hard termination can also leave a native Git index.lock; investigate native
Git locks before resuming rather than deleting them while Git may still run.
Use actual directory paths for plans/logs: directory symlink/junction ancestry
is refused for these owned outputs so an external-looking path cannot redirect
writes into the repository. Ordinary OneDrive cloud placeholders are not banned.
Filters changing shard bytes (including LFS pointers) fail verification;
use a repository configuration approved for preserving these file bytes.

Runtime limits: 10,000 shard/manifest files, 256 MiB per shard, 8 MiB per JSON
document, 2 MiB captured output per native stream, 100 log files retained.
Child executables: installed git.exe and its configured approved helpers/hooks;
Windows taskkill.exe only for the process tree owned by this invocation on
timeout/cancellation. No network contact occurs until an authorized Push.

## Deployment and help

Deploy `sendshards.ps1` alone with its declared Windows/Git dependencies, approved credentials and runtime plan/input data. It does not require Embelish, XML, Markdown or agent instructions to plan/resume a transfer. Optional XML supplies detailed native help; without it, `Get-Help` and `-?` fall back to generated syntax. To retain detailed help, also copy matching `embelish/` and `docs/` directories together so the native `.EXTERNALHELP` directive resolves `../docs/sendshards/sendshards-help.xml`. Keep Markdown and native XML help synchronized when documentation changes.

## Related tool

[Embelish](../embelish/README.md) creates the encrypted base64 shards. Sendshards verifies transferred bytes but does not decrypt or establish the archive's authenticity.

## Vendor references

- https://git-scm.com/docs/git-push
- https://git-scm.com/docs/git-ls-remote
- https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/taskkill
