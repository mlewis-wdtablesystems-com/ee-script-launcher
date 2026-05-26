# WDTS PLM / Altium automation workspace

PowerShell scripts and a web **launcher** for Arena PLM and Altium Designer workflows (BOM compare, file sync, item duplicate, supplier BOM export, and related tasks).

## Layout

| Path | Purpose |
|------|---------|
| `scripts/` | PowerShell automation scripts |
| `scripts/launcher/` | FastAPI launcher UI (`python app.py` or `launcher.bat`) |
| `canvases/` | Optional Cursor canvas artifacts |
| `wdts-altium-mcp/` | Separate git repo — Altium MCP server (not tracked in this repo) |

## Prerequisites

- Windows with PowerShell 5.1+
- Arena credentials via environment variables: `ARENA_EMAIL`, `ARENA_PASSWORD`, `ARENA_WORKSPACE_ID`
- Python 3.10+ for the launcher (`pip install -r scripts/launcher/requirements.txt`)

## Launcher

```text
cd scripts\launcher
.\launcher.bat
```

Open the URL shown in the console (default http://127.0.0.1:8765).

Starting the launcher again stops any prior instance on that port (use `run.ps1 -NoForce` to keep an existing server).

## Version control

This repository uses **git**. Make focused commits with clear messages after each completed change. Do not commit secrets, `.env` files, virtualenvs, logs, or large generated CSV/xlsx exports.
