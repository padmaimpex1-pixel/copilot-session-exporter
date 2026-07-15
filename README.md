# copilot-session-exporter

Export all [GitHub Copilot CLI](https://githubnext.com/projects/copilot-cli) session prompts and assistant responses into a formatted Excel workbook (`.xlsx`).

## Requirements

- Python 3.10+
- [openpyxl](https://openpyxl.readthedocs.io/)

```bash
pip install openpyxl
```

## Usage

```bash
python export_sessions.py                    # outputs copilot_sessions.xlsx in current dir
python export_sessions.py my_export.xlsx     # custom output path
```

## Output

The workbook contains two sheets:

### Messages
Flat log of every message across all sessions:

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

Rows are colour-coded: **blue** for user prompts, **green** for assistant replies. Auto-filter and frozen header row are enabled.

### Summary
One row per session with interaction and message counts.

## How it works

The script reads `~/.copilot/session-state/*/events.jsonl` — the event log that Copilot CLI writes for every session. It extracts `user.message` and `assistant.message` events, groups them by interaction ID, and writes everything to Excel.
