# Mercari JP Search & GUI

A Python application to search Mercari Japan products with a Web GUI. Supports both simple and bulk keyword searches with powerful filtering options.

## Features

- **Simple Search** — Search by keyword with price range, status, condition filters
- **Bulk Search** — Upload a `.txt`, `.csv`, or `.xlsx` file (or paste keywords) and search multiple keywords at once
- **Filters** — Price range, item status, condition, seller review count (bulk search)
- **Shop Item Detection** — Automatically detects `/shops/product/` vs `/item/` URLs
- **Description Viewer** — View and translate item descriptions via modal
- **CSV Export** — Download bulk search results as a CSV file
- **Japanese Translation** — Product names and descriptions automatically translated to English

## Requirements

- Python 3.11+
- Dependencies listed in `pyproject.toml`

## Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/rafayqazi/mercari-jp-mcp.git
   cd mercari-jp-mcp
   ```

2. Create and activate a virtual environment:
   ```bash
   uv venv
   source .venv/bin/activate  # Windows: .venv\Scripts\activate
   uv pip install -e .
   ```

## Usage

### Web GUI (Flask Application — Recommended)

Run the built-in Flask web server:

```bash
uv run mercari_gui.py
```

Then open **http://127.0.0.1:5000** in your browser.

**Windows users** can also double-click `start_gui.bat` to launch directly.

The GUI has two tabs:

- **Simple Search** — Single keyword with price range, exclude keywords, status/condition filters
- **Bulk Search** — Paste multiple keywords (one per line) or upload a `.txt` / `.csv` / `.xlsx` file. Filter by status, condition, and seller review count (min/max). Results are grouped by keyword and can be exported as CSV.

Click **Description** on any result to view item details and auto-translate the description from Japanese to English. Click **View** to open the item on Mercari Japan.

### MCP Server (for Claude Desktop)

```bash
uv run server.py
```

Configure in `claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "mercari": {
      "command": "uv",
      "args": [
        "--directory",
        "C:\\PATH\\TO\\mercari-jp-mcp",
        "run",
        "server.py"
      ]
    }
  }
}
```

## Security

- No API keys, tokens, or credentials are stored in this repository
- All search requests go directly to Mercari Japan's public API
- The `.gitignore` excludes virtual environments, cache files, and local config
- Environment files (`.env`, `.env.*`) are excluded — never commit secrets

## How It Works

The app uses the [mercari](https://github.com/marvinody/mercari/) library (MIT license) to interface with Mercari Japan's public API. Search results are parsed, filtered, and optionally translated via Google Translate API.

## Acknowledgments

- [marvinody/mercari](https://github.com/marvinody/mercari/) — Mercari API client
- [jlowin/fastmcp](https://github.com/jlowin/fastmcp) — MCP framework
