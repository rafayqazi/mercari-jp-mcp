# Mercari JP Search + Yahoo Auctions GUI

A Python application to search Mercari Japan and Yahoo Auctions products with a Web GUI. Supports simple, bulk, and combined cross-platform searches with powerful filtering options.

## Features

- **Mercari Simple Search** — Single keyword with price range, exclude keywords, status/condition filters
- **Mercari Bulk Search** — Upload a `.txt`, `.csv`, or `.xlsx` file (or paste keywords) and search multiple Mercari keywords at once
- **Yahoo Auctions Simple Search** — Search Yahoo Auctions with price range, status (live/sold), condition, sort, and BIN-only filters
- **Yahoo Auctions Bulk Search** — Multi-keyword Yahoo Auctions search with CSV export
- **Combined Mercari + Yahoo Bulk Search** — Search both platforms simultaneously with per-keyword results shown side-by-side
- **Filters** — Price range, item status, condition, seller review count (Mercari bulk), sort order, BIN-only (Yahoo)
- **Shop Item Detection** — Automatically detects `/shops/product/` vs `/item/` URLs
- **Description Viewer** — View item descriptions with auto-translate from Japanese to English
- **CSV Export** — Download search results as CSV for all search types
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

The GUI has five tabs:

- **Simple Search** — Single keyword with price range, exclude keywords, status/condition filters
- **Mercari Bulk Search** — Paste multiple keywords (one per line) or upload a `.txt` / `.csv` / `.xlsx` file. Filter by status, condition, and seller review count (min/max). Results are grouped by keyword and can be exported as CSV.
- **Yahoo Auctions** — Single keyword search on Yahoo Auctions with price range, status (live/sold), condition, sort order, and BIN-only filter
- **Yahoo Bulk** — Multi-keyword Yahoo Auctions search with file upload support and CSV export
- **Mercari + Yahoo Bulk** — Search both platforms at once. Results are displayed side-by-side per keyword with Mercari on the left and Yahoo on the right.

Click **Description** (Mercari) or **Detail** (Yahoo) on any result to view item details with auto-translate from Japanese to English. Click **View** to open the item on the respective platform.

### MCP Server (for Claude Desktop)

The MCP server exposes a `search_mercari_jp` tool that returns formatted markdown results:

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

The server automatically limits API pagination (`max_items`) for faster responses and returns results as markdown with images and links.

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
