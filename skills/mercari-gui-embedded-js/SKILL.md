---
name: mercari-gui-embedded-js
description: >-
  Rules for editing JavaScript inside Python triple-quoted HTML in mercari_gui.py
  (HTML_PAGE). Prevents silent JS parse errors that break tabs, buttons, and onclick
  handlers. Use when changing mercari_gui.py UI, bulk streaming, CSV download, tabs,
  or any embedded <script> in this repo.
---

# Mercari GUI — embedded JavaScript in Python

The Flask UI lives in `mercari_gui.py` as `HTML_PAGE = u'''...'''`. All client logic is one inline `<script>`. **One syntax error anywhere in that script prevents the entire block from parsing** — tabs, search buttons, and bulk streaming all stop working with no obvious server error.

## Golden rule: Python string vs JavaScript string

Inside `HTML_PAGE`, you are writing **Python source** that must emit **valid JavaScript**.

| You want in the browser (JS) | Write in mercari_gui.py (Python) |
|-----------------------------|----------------------------------|
| Newline character in JS: `'\n'` | `'\\n'` |
| Split on blank line: `'\n\n'` | `'\\n\\n'` |
| Backslash in JS string | `'\\\\'` (only when needed) |

**Never** use bare `'\n'` or `'\n\n'` inside the HTML/JS portion of `HTML_PAGE`. Python treats `\n` as a real line break in the string, which **splits JS string literals across physical lines** and causes `SyntaxError: Invalid or unexpected token`.

### Wrong → broken output

Python source (wrong):

```javascript
text.split('\n')
data.keywords.join('\n')
buffer.split('\n\n')
let csv = 'Header\n';
```

Rendered JS (broken — string literal split across lines):

```javascript
text.split('
')
```

### Right

```javascript
text.split('\\n')
data.keywords.join('\\n')
buffer.split('\\n\\n')
let csv = 'Header\\n';
csv += row + '"\\n';
```

Plain Python code **outside** `HTML_PAGE` (e.g. `joined = '\n'.join(texts)`) uses normal Python `'\n'` — do not double-escape there.

## Other escaping in this file

- **`&` in JS strings inside HTML**: Use HTML entities in the embedded HTML/JS text when the character would break HTML parsing, e.g. `Loading &amp; translating...` not raw `&` inside attribute-like contexts.
- **Item IDs in dynamic HTML**: Prefer `data-item-id="..."` + delegated click handlers (see bottom of script). Avoid fragile `onclick="fetchDescription('...')"` with nested quotes.
- **Tabs**: Use `data-tab` + `switchTab(name)`; tab panel ids are `tab` + PascalCase suffix (`bulk` → `tabBulk`, `yahooBulk` → `tabYahooBulk`).

## Architecture (do not break)

- **Simple / Yahoo single search**: `fetch` + JSON, `renderResults` / `renderYahooResults`.
- **Bulk (Mercari, Yahoo, Combined)**: POST + **SSE-style** stream via `streamBulkSearch()`; lines `data: {...}\n\n`. Append with `appendBulkResult` / `appendYahooBulkResult` / `appendCombinedResult`; do not replace whole DOM on each chunk unless intentional.
- **CSV rows**: End each line with `'"\\n'` in Python so JS emits `\n` in the CSV blob.

## After every edit to `<script>` or `HTML_PAGE`

Run validation (extract **runtime** string, not raw file slice):

```bash
python .cursor/skills/mercari-gui-embedded-js/scripts/validate_embedded_js.py
```

Must exit 0. If Node is missing, at least import the module and inspect `_runtime.js` manually.

## Checklist before finishing

- [ ] No new `'\n'` / `'\n\n'` inside `HTML_PAGE` JavaScript (only `'\\n'` / `'\\n\\n'`).
- [ ] `node --check` passes on extracted runtime JS.
- [ ] Tab ids match `switchTab` naming (`tab` + `name[0].toUpperCase() + name.slice(1)`).
- [ ] Restart Flask and hard-refresh (Ctrl+F5) when testing the UI.

## Reference

- Main file: `mercari_gui.py` (`HTML_PAGE`, routes under `/api/`).
- Historical bug: bulk streaming change used unescaped `'\n'` → entire UI frozen.
