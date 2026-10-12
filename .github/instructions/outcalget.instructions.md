---
applyTo: "outcalget/**,docs/outcalget/**"
---

# outcalget instructions for GitHub Copilot

Read `outcalget/AGENTS.md` and follow its full scoped contracts. Read `docs/outcalget/README.md` and the usage/operations/implementation/verification documents relevant to the change.

Preserve the native Outlook exporter, 30-local-date/full-detail semantics, legacy configurations, calendar identity checks, synchronization and exact-process ownership. Preserve locks, output claims, bounded worker budgets, recovery state, prior-output protection and redacted diagnostics. Do not force-close Outlook or infer server freshness from group completion.

Keep `outcalget.ps1` implementation-focused and preserve its native external-help directive. Update corresponding Markdown and `docs/outcalget/outcalget-help.xml` together when public help changes. Verify help from an unrelated working directory.

Do not execute `update-calendar.ps1` as a development check. For documentation extraction, compare executable tokens, preserve encodings/native source strings, and check help and version without reaching business work. Report target Windows/Outlook/Exchange/Scheduler checks separately.
