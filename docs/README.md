# Documentation

The documentation and XML help files are optional for runtime execution. The exporter, Embelish and Sendshards can each be deployed as their `.ps1` alone, with the declared platform/application dependencies and runtime input/configuration. The existing calendar Git wrapper also needs `outcalget.ps1` beside it. XML adds detailed native help; without it, PowerShell falls back to generated parameter syntax.

- [PowerShell implementation guidelines](Powershell%20guidelines.md)
- [outcalget documentation](outcalget/README.md): exporter usage, scheduling, configuration, recovery, implementation notes and verification
- [Embelish documentation](embelish/README.md): interactive encrypted archive/shard creation and restoration
- [Sendshards documentation](sendshards/README.md): local transfer planning and verified resumable Git pushes
- [Agent configuration](agent-configuration.md): shared instructions for GitHub Copilot, Claude Code and Codex

## Script coverage

| Script | Documentation | Scoped agent instructions |
|---|---|---|
| `outcalget/outcalget.ps1` | [Exporter](outcalget/README.md) | `outcalget/AGENTS.md` |
| `outcalget/update-calendar.ps1` | [Calendar Git helper](outcalget/update-calendar.md) | `outcalget/AGENTS.md` |
| `embelish/embelish.ps1` | [Archive and restore](embelish/README.md) | `embelish/AGENTS.md` |
| `embelish/sendshards.ps1` | [Shard transfer](sendshards/README.md) | `embelish/AGENTS.md` |

These are all four PowerShell scripts currently in the repository. Each retains its existing executable implementation; long help, operation notes, function rationale and verification evidence are organized under `docs/`.
