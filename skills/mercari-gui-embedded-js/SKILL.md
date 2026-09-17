---
name: mercari-gui-embedded-js
description: >-
  Complete reference for working on mercari_gui.py — the Flask GUI for
  Mercari JP + Yahoo Auctions search. Covers embedded JS escaping rules,
  full project architecture, API routes, the Mercari library internals,
  known bugs & fixes, and how to keep this skill updated after changes.
  Use whenever editing mercari_gui.py UI, bulk streaming, CSV download,
  tabs, API routes, or the mercari/yahoo_auctions/ebay_search libraries.
---

# Mercari GUI — Full Project Reference

## Project Layout

```
mercari-jp-mcp/
├── mercari_gui.py                  # Main file: Flask app + entire HTML/JS UI
├── server.py                       # MCP server (separate from GUI)
├── start_gui.bat                   # Launcher: starts Flask on port 5000
├── yahoo_auctions/                 # Yahoo Auctions search library
├── ebay_search/                    # eBay search + OAuth library
├── temp_download/
│   └── mercari-2.2.1/
│       └── mercari/
│           ├── mercari.py          # Core library: Item class, search(), getItemInfo()
│           ├── MercariItemFull.py  # Full item detail model (used by getItemInfo)
│           └── DpopUtils.py        # DPOP auth header generator
└── skills/
    └── mercari-gui-embedded-js/
        └── SKILL.md                # This file
```

---

## mercari_gui.py — Structure

`mercari_gui.py` is a **single file** Flask app (~3600 lines). It contains:

1. **Python imports + Flask setup** (top ~20 lines)
2. **`HTML_PAGE = u'''...'''`** — the entire HTML + CSS + JavaScript UI as a Python triple-quoted string
3. **Flask API route handlers** — all `@app.route('/api/...')` functions below `HTML_PAGE`

### UI Tabs (in order)
| Tab name | `switchTab()` id | Description |
|---|---|---|
| Mercari Search | `mercari` | Single keyword Mercari search |
| Yahoo Auctions | `yahoo` | Single keyword Yahoo search |
| Mercari Bulk | `mercariBulk` | Bulk Mercari by keywords |
| Yahoo Bulk | `yahooBulk` | Bulk Yahoo by keywords |
| Mercari+Yahoo | `combinedBulk` | Combined Mercari+Yahoo+eBay ref |
| Stock Check | `stockCheck` | Check status of pasted links |
| eBay Search | `ebay` | Single eBay search |
| eBay Bulk | `ebayBulk` | Bulk eBay by keywords |

Tab panel IDs = `tab` + PascalCase name, e.g. `tabMercariBulk`, `tabCombinedBulk`.

---

## Flask API Routes

| Route | Method | Stream? | Description |
|---|---|---|---|
| `/` | GET | No | Serves `HTML_PAGE` |
| `/api/search` | POST | No | Single Mercari search |
| `/api/yahoo-search` | POST | No | Single Yahoo search |
| `/api/bulk-search` | POST | SSE | Mercari bulk (per keyword) |
| `/api/yahoo-bulk-search` | POST | SSE | Yahoo bulk (per keyword) |
| `/api/combined-bulk-search` | POST | SSE | Mercari+Yahoo+eBay combined |
| `/api/check-stock` | POST | SSE | Stock check for pasted links |
| `/api/item-description` | POST | No | Fetch Mercari item description |
| `/api/yahoo-item-detail` | POST | No | Fetch Yahoo item detail |
| `/api/ebay-search` | POST | No | Single eBay search |
| `/api/ebay-bulk-search` | POST | SSE | eBay bulk search |
| `/api/upload-keywords` | POST | No | Parse uploaded XLSX/CSV file |

### SSE Streaming Pattern
All bulk/stock routes use Server-Sent Events:
- Python yields: `'data: ' + json.dumps({...}) + '\n\n'`
- JS reads with `streamBulkSearch(url, body, callback)` function
- Each chunk calls the callback with parsed data
- Final chunk: `{'complete': True, 'keywords': [...]}`

---

## Mercari Library Internals (`mercari/mercari.py`)

### `Item` class — search result object

```python
class Item:
    id               # str  -- item ID
    productURL       # str  -- https://jp.mercari.com/item/{id}  (regular)
                     #         https://jp.mercari.com/shops/product/{id}  (shop)
    imageURL         # str  -- first thumbnail URL
    productName      # str  -- item name (Japanese)
    price            # int  -- price in JPY
    status           # str  -- 'ITEM_STATUS_ON_SALE' | 'ITEM_STATUS_SOLD_OUT' | ...
    isShopItem       # bool -- True if Mercari Shops (BEYOND) item
    itemConditionId  # str  -- '1'..'6'
    auction          # ItemAuction | None
```

### CRITICAL: Shop Item Detection (fixed 2026-08)

The Mercari API does **NOT** return `isShopItem` or `is_shop_item` fields.
Shop items are identified by:

| Field | Regular item | Shop item |
|---|---|---|
| `itemType` | `"ITEM_TYPE_MERCARI"` | `"ITEM_TYPE_BEYOND"` |
| `shop` | `null` | `{"id": "..."}` |
| ID format | starts with `m` + digits | alphanumeric base62 |

**The fix** (applied in `mercari.py` `Item.__init__`):
```python
item_type = kwargs.get('itemType', '')
shop_data = kwargs.get('shop', None)
raw = kwargs.get('isShopItem', kwargs.get('is_shop_item', False))
legacy_flag = str(raw).lower() in ('true', '1') if not isinstance(raw, bool) else raw
self.isShopItem = legacy_flag or item_type == 'ITEM_TYPE_BEYOND' or shop_data is not None
# Auto-correct productURL for shop items
if self.isShopItem:
    self.productURL = "https://jp.mercari.com/shops/product/{}".format(kwargs['id'])
```

Before this fix: all shop items got `/item/` URLs → 404 on View click.
After this fix: `item.productURL` and `item.isShopItem` are always correct.

The redundant `isShopItem` checks in `mercari_gui.py` (lines ~2356, ~2572, ~2897)
still work correctly as a safety net.

### Mercari URL formats
```
Regular item:   https://jp.mercari.com/item/{id}
Shop product:   https://jp.mercari.com/shops/product/{id}
```

### `getItemInfo(itemID)` — full item detail
- Calls `/items/get` API
- Returns `MercariItemFull` with seller info, full description, etc.
- Used for review count filtering and description fetch
- **Do NOT call in bulk loops without explicit need** — slow (1 request per item)

### `getShopProductInfo(itemID)` — shop product detail
- Calls `/v1/marketplaces/shops/products/{id}` API
- Returns `ShopProductInfo` with `.name`, `.price`, `.status`, `.thumbnail`

---

## Stock Check — Supported Link Types

The `/api/check-stock` route parses these URL patterns:

| Platform | URL pattern | Type key |
|---|---|---|
| Yahoo Auctions | `auctions.yahoo.co.jp/auction/{id}` | `yahoo` |
| Yahoo Auctions (page.) | `page.auctions.yahoo.co.jp/auction/{id}` | `yahoo` |
| PayPay Fleamarket | `paypayfleamarket.yahoo.co.jp/item/{id}` | `paypay` |
| Mercari Shops | `mercari.com/shops/product/{id}` | `mercari_shop` |
| Mercari | `mercari.com/item/{id}` | `mercari` |
| HardOff NetMall | `netmall.hardoff.co.jp/product/{id}` | `netmall` |
| Rakuten | `item.rakuten.co.jp/{shop}/{id}` | `rakuten` |
| Yahoo Shopping | `store.shopping.yahoo.co.jp/{shop}/{id}.html` | `yahoo_shopping` |

---

## Embedded JavaScript Escaping Rules

The Flask UI lives in `HTML_PAGE = u'''...'''`. All client logic is one inline `<script>`.
**One syntax error anywhere in that script prevents the entire block from parsing** —
tabs, search buttons, and bulk streaming all stop working with no obvious server error.

### Golden Rule: Python string vs JavaScript string

Inside `HTML_PAGE`, you are writing **Python source** that must emit **valid JavaScript**.

| You want in the browser (JS) | Write in mercari_gui.py (Python) |
|-----------------------------|----------------------------------|
| Newline `'\n'` | `'\\n'` |
| Double newline `'\n\n'` | `'\\n\\n'` |
| Backslash in JS string | `'\\\\'` |

**Never** use bare `'\n'` inside the JS portion of `HTML_PAGE`. Python treats `\n` as a real
line break, which **splits JS string literals** and causes `SyntaxError`.

### Wrong (broken)
```python
# Inside HTML_PAGE:
text.split('\n')      # splits JS string literal across lines
csv += row + '"\n';   # broken
buffer.split('\n\n')  # broken
```

### Right (correct)
```python
# Inside HTML_PAGE:
text.split('\\n')
csv += row + '"\\n';
buffer.split('\\n\\n')
```

Plain Python code **outside** `HTML_PAGE` uses normal `'\n'` — do not double-escape there.

### Other escaping
- **`&` in embedded HTML/JS**: Use HTML entities: `&amp;`, `&rarr;`, `&yen;` etc.
- **Item IDs in dynamic HTML**: Use `data-item-id="..."` + delegated handlers, not `onclick="fn('...')"`

---

## Architecture Rules (do not break)

- **Simple / Yahoo single search**: `fetch` + JSON -> `renderResults()` / `renderYahooResults()`
- **Bulk streams**: POST + SSE via `streamBulkSearch(url, body, callback)`; append with
  `appendBulkResult` / `appendYahooBulkResult` / `appendCombinedBulkResult`
- **Retry buttons**: Store config in `window._retryBulkConfigs[tabName]` before streaming starts
- **CSV rows**: End each line with `'"\\n'` in Python so JS emits `\n` in the CSV blob
- **Dedup**: Each bulk tab uses a `Set` of keys (`m|kw|url|price` pattern) to skip duplicates on retry
- **eBay credentials**: Read from `getEbayCredentials()` stored in `localStorage`
- **View links**: Always use `item.url` directly from API response — never reconstruct from `item.id`

---

## After Every Edit to `<script>` or `HTML_PAGE`

Run validation (extract **runtime** string, not raw file slice):

```bash
python .cursor/skills/mercari-gui-embedded-js/scripts/validate_embedded_js.py
```

Must exit 0. If Node is missing, at least import the module and inspect `_runtime.js` manually.

### Checklist before finishing
- [ ] No bare `'\n'` / `'\n\n'` inside `HTML_PAGE` JavaScript — only `'\\n'` / `'\\n\\n'`
- [ ] `node --check` passes on extracted runtime JS
- [ ] Tab ids match `switchTab` naming (`tab` + PascalCase)
- [ ] `item.url` is used directly for View links — never reconstruct from `item.id` alone
- [ ] Restart Flask and hard-refresh (Ctrl+F5) when testing the UI

---

## Keeping This Skill Updated

After making significant changes (new routes, bug fixes, new URL patterns, new library
behaviour), update this SKILL.md so future AI sessions start with correct context.

### Diagnostic commands

Snapshot current route list:
```bash
python check_server.py
```
Or manually:
```bash
python -c "
import ast
src = open('mercari_gui.py', encoding='utf-8').read()
for node in ast.walk(ast.parse(src)):
    if isinstance(node, ast.FunctionDef):
        for dec in node.decorator_list:
            if isinstance(dec, ast.Call) and hasattr(dec.func, 'attr') and dec.func.attr == 'route':
                print(dec.args[0].value if dec.args else '?', '->', node.name)
"
```

Verify shop item detection fix is intact in the library:
```bash
python -c "
import sys; sys.path.insert(0,'temp_download/mercari-2.2.1')
from mercari.mercari import Item
i = Item(id='X',name='t',price=1,status='s',created=0,updated=0,
         thumbnails=['u'],itemType='ITEM_TYPE_BEYOND',shop={'id':'y'},auction=None)
assert i.isShopItem and 'shops/product' in i.productURL
print('OK:', i.productURL)
"
```

### What to document when you make changes
1. **New route added** → add to the API Routes table above
2. **New link type in stock check** → add to the Supported Link Types table
3. **Library bug fixed** → add a dated entry to the Known Bugs & Fixes Log below
4. **New JS pattern/rule** → add to Architecture Rules or Escaping Rules
5. **New known bug** → add to Known Bugs & Fixes Log with symptom, cause, and fix

---

## Known Bugs & Fixes Log

### 2025 — Bulk streaming broken (newline escaping)
**Symptom**: Entire UI frozen/broken after adding bulk streaming.  
**Cause**: Used `'\n'` inside `HTML_PAGE` JS string — split literal across lines.  
**Fix**: Always use `'\\n'` inside `HTML_PAGE`.

### 2026-08 — Shop product View button opens /item/ (404)
**Symptom**: Mercari Shops products (in combined/bulk search) show correct image + title
but clicking View opens `https://jp.mercari.com/item/{id}` which returns a 404 page.  
**Cause**: `Item.isShopItem` was always `False` because the Mercari API returns `itemType`
and `shop` fields — not `isShopItem`/`is_shop_item`. The library never correctly detected
shop items, so `productURL` was always the `/item/` path.  
**Fix**: Updated `Item.__init__` in `temp_download/mercari-2.2.1/mercari/mercari.py` to
detect shop items via `itemType == 'ITEM_TYPE_BEYOND'` or `shop != None`. Also
auto-corrects `productURL` to `/shops/product/` for shop items at the library level.
All existing `isShopItem` checks in `mercari_gui.py` now work correctly.

---

## Reference

- Main file: `mercari_gui.py` (`HTML_PAGE` + all `/api/` routes)
- Mercari library: `temp_download/mercari-2.2.1/mercari/mercari.py`
- Historical bug: bulk streaming change used unescaped `'\n'` -> entire UI frozen
