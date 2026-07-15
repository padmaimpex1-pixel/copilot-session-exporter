# copilot-session-exporter

Two scripts to export and query [GitHub Copilot CLI](https://githubnext.com/projects/copilot-cli) session history.

## Requirements

- Python 3.10+
- [openpyxl](https://openpyxl.readthedocs.io/) (export + query)
- [pandas](https://pandas.pydata.org/) (query only)

```bash
pip install openpyxl pandas
```

---

## export_sessions.py — Export to Excel

Exports all Copilot CLI session prompts and assistant responses from
`~/.copilot/session-state/` into a formatted `.xlsx` workbook.
Sessions are ordered newest-first.

```bash
python export_sessions.py                    # outputs copilot_sessions.xlsx in cwd
python export_sessions.py my_export.xlsx     # custom output path
```

### Output sheets

**Messages** — flat log of every message:

| Column | Description |
|---|---|
| Session ID | UUID of the Copilot session |
| Session Start (UTC) | When the session started |
| Workspace | Folder path open in VS Code |
| Interaction # | Sequential number within the session |
| Turn # | Assistant turn index within an interaction |
| Timestamp (UTC) | Message timestamp |
| Role | `user` or `assistant (model-name)` |
| Message | Full prompt or response text |

Rows are colour-coded: **blue** for user prompts, **green** for assistant replies.

**Summary** — one row per session with interaction and message counts.

---

## query_sessions.py — Query the Excel file

CLI tool for searching and filtering the exported workbook.

```
usage: query_sessions [--file FILE] {sessions,show,search,stats} ...
```

### Subcommands

```bash
# List all sessions
python query_sessions.py sessions
python query_sessions.py sessions --since 2026-07-01
python query_sessions.py sessions --workspace Myapp

# Show all messages in a session (prefix match on session ID)
python query_sessions.py show 34b88269
python query_sessions.py show 34b88269 --limit 10
python query_sessions.py show 34b88269 --output session_export.xlsx

# Full-text search (supports regex)
python query_sessions.py search "angular"
python query_sessions.py search "error|exception" --role assistant
python query_sessions.py search "storefront" --since 2026-06-01 --limit 20
python query_sessions.py search "deploy" --output deploy_mentions.xlsx

# Aggregate statistics
python query_sessions.py stats
python query_sessions.py stats --since 2026-07-11
python query_sessions.py stats --workspace bitwarden
```

### Flags (available on all subcommands)

| Flag | Description |
|---|---|
| `--since YYYY-MM-DD` | Include rows on or after this date |
| `--until YYYY-MM-DD` | Include rows on or before this date |
| `--workspace TEXT` | Filter by workspace path substring (case-insensitive) |
| `--limit N` | Cap output rows |
| `--output FILE` | Save matched rows to a new xlsx instead of printing |
| `--file FILE` | Path to xlsx (default: `copilot_sessions.xlsx` in cwd) |

### stats output example

```
=== Stats ===

  Sessions         : 37
  Interactions     : 401
  Total messages   : 3238
  User messages    : 406
  Assistant msgs   : 2832

  Model usage:
    gpt-5.3-codex                       1067
    claude-sonnet-4.6                    778

  Messages per day (last 10):
    2026-07-15                                 13
    2026-07-14  ######                         140
    2026-07-13  #################              356
```
