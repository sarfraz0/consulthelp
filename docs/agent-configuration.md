# Agent configuration

[Documentation index](README.md)

The repository uses native instruction files for GitHub Copilot, Claude Code and Codex. No additional runtime, plugin, secrets or broad tool permissions are needed for these instructions. Each agent reads the same engineering policy and applicable tool documentation. Coverage includes all four scripts in `outcalget/` and `embelish/`, with separate documentation for Sendshards.

| Agent | Repository entry point | outcalget scope |
|---|---|---|
| Codex | Root `AGENTS.md` | `outcalget/AGENTS.md`, discovered for work in that directory |
| Claude Code | Root `CLAUDE.md` imports `@AGENTS.md` | `outcalget/CLAUDE.md` imports the local `@AGENTS.md` when Claude works in that subtree; root guidance also directs it to read scoped instructions |
| GitHub Copilot | `.github/copilot-instructions.md` | `.github/instructions/outcalget.instructions.md`, with `applyTo: "outcalget/**,docs/outcalget/**"` |

`AGENTS.md` and `outcalget/AGENTS.md` are the shared workflow and behavior instructions. The Claude files import them using Claude Code's native Markdown import syntax. Copilot's recognized instruction files provide the essential rules and explicitly direct it to read the shared files; ordinary Markdown references are not claimed to be automatic imports. Clients with repository/path instruction support disabled must enable it or explicitly provide these files. Copilot products differ in their support for path-specific instructions; the repository-wide file retains the key rules.

For both `embelish/embelish.ps1` and `embelish/sendshards.ps1`, Codex reads `embelish/AGENTS.md` and Claude Code imports that file through `embelish/CLAUDE.md`. Copilot uses `.github/instructions/embelish.instructions.md` with `applyTo: "embelish/**,docs/embelish/**,docs/sendshards/**"`. This shared directory scope contains distinct archive and transfer contracts; read the relevant tool's documentation before editing.

## Working convention

Read the requested scope, inspect the working tree, then read the shared guidelines and scoped documentation before editing. Preserve the existing runtime and contracts, make a focused change, run proportionate checks, and update matching documentation. Report current verification separately from historical records and unavailable target checks.

These files are instructions, not security boundaries or grants to run corporate workflows. Existing client permissions continue to govern tools. No auto-approve modes, policy bypasses, hooks, external integrations or credential settings are configured.

## Maintaining the configuration

Update shared rules in `AGENTS.md` and the applicable scoped file. Change the detailed engineering policy in [`Powershell guidelines.md`](Powershell%20guidelines.md). Keep the short Copilot summaries consistent with those rules. Claude imports and Codex discovery should continue to reference the same shared files.

Use documentation paths relative to the repository root when instructions name them. A task on files outside `outcalget/`, such as `docs/outcalget/usage.md`, must still read `outcalget/AGENTS.md` explicitly; filesystem discovery alone does not apply a sibling directory's instructions.

Likewise, a task on `docs/embelish/` or `docs/sendshards/` must read `embelish/AGENTS.md` explicitly. The shared engineering defaults do not replace each tool's established runtime, native interfaces, CLI, diagnostic and side-effect contracts.

Configuration validation checks the recognized filenames, Claude import targets, Copilot front matter and local documentation links. No live GitHub Copilot, Claude Code or Codex client session was exercised to certify instruction loading.
