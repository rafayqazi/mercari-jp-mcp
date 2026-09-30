# Yahoo! Auctions Scraping & Architecture Guide

> **For Future AI Agents & Developers:**  
> This document details the exact architecture, data flows, endpoint specifications, and JSON structures used for Yahoo! Auctions integration in this codebase. If Yahoo! Auctions updates its UI or API in the future, follow this guide to understand the system and apply fixes quickly.

---

## 1. High-Level Architecture Overview

Yahoo! Auctions (`https://auctions.yahoo.co.jp/`) runs on a modern **Next.js** frontend.

- **No HTML DOM Item Cards:** Search result cards are **not** rendered as static server-side HTML tags with persistent CSS classes (such as old `li.Product` or `Product__titleLink`). CSS class names are hashed modules (e.g. `gv-Carousel__item--SaTobYP4YqSQeAzPrIGF`).
- **Single Source of Truth:** All listings and item details are embedded directly inside the document as a JSON payload within:
  ```html
  <script id="__NEXT_DATA__" type="application/json">
    { "props": { ... } }
  </script>
  ```
- **Dual Platform Listings:** Yahoo! Auctions search returns both standard auction listings and **Yahoo! Flea Market / PayPay Fleamarket** items (`isFleamarketItem: true`, item IDs typically starting with `z`). These two platforms have different detail page URLs and JSON state schemas.

---

## 2. Directory & Module Responsibilities

All Yahoo! Auctions logic is located inside [`temp_download/yahoo_auctions/`](file:///c:/Users/AR/Desktop/mercari-jp-mcp/temp_download/yahoo_auctions/):

```
temp_download/yahoo_auctions/
├── __init__.py       # Package exports (search_yahoo, get_item_detail, parse_search_results, etc.)
├── models.py         # YahooAuctionItem dataclass definition
├── client.py         # Throttled HTTP client, anti-bot session warmup, soft-404 detection
├── search.py         # Search URL builders (live vs closed), Next.js JSON search parser, legacy fallback
└── item.py           # Item detail scraper (Yahoo Auctions & PayPay Fleamarket dual resolution)
```

And integrated into the Flask application:
- [`mercari_gui.py`](file:///c:/Users/AR/Desktop/mercari-jp-mcp/mercari_gui.py):
  - `/api/yahoo-search`: Single keyword search endpoint
  - `/api/yahoo-bulk-search`: SSE streaming multi-keyword search
  - `/api/combined-bulk-search`: Side-by-side Mercari + Yahoo bulk search
  - `/api/yahoo-item/<item_id>`: Detailed modal inspection
  - `/api/check-stock`: Availability and stock checker for URLs

---

## 3. Endpoints & URL Rules

| Operation | Base URL | Key Parameters | Notes |
| :--- | :--- | :--- | :--- |
| **Live Search** | `https://auctions.yahoo.co.jp/search/search` | `p` (query), `auccat=new`, `b` (offset = `(page-1)*limit+1`), `n` (limit), `s` (sort), `min`, `max`, `istatus` | Returns active listings |
| **Sold / Closed Search** | `https://auctions.yahoo.co.jp/closedsearch/closedsearch` | `p` (query), `b` (offset), `n` (limit), `s` (sort), `min`, `max`, `istatus` | **Do not** pass `auccat=closed` to `search/search`. Must use dedicated `closedsearch` route. |
| **Auction Detail** | `https://page.auctions.yahoo.co.jp/auction/{item_id}` | Path variable `item_id` | Standard auctions |
| **Flea Market Detail** | `https://paypayfleamarket.yahoo.co.jp/item/{item_id}` | Path variable `item_id` | PayPay Fleamarket items (`z...`) |

### Sort Parameter Mapping (`s` parameter)
- `new` &rarr; `-new` (Newest listed)
- `end` &rarr; `end` (Ending soonest)
- `price` &rarr; `price` (Price: Low to High)
- `-price` &rarr; `-price` (Price: High to Low)
- `bid` &rarr; `-bid` (Most bids)

---

## 4. Search Results Parsing Specification

### Step 1: Extract `__NEXT_DATA__`
Use regex to isolate the JSON string:
```python
m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.DOTALL)
nd = json.loads(m.group(1))
```

### Step 2: Navigate Search Items
The listings array is located at:
```python
props = nd.get('props', {})
page_props = props.get('pageProps', {})
init = page_props.get('initialState') or props.get('initialState', {})
search = init.get('search', {})
items_state = search.get('items', {})

# Primary listing array:
raw_items = items_state.get('listing', {}).get('items', [])

# Fallback listing array if empty:
if not raw_items:
    raw_items = items_state.get('featuredListing', {}).get('items', [])
```

### Step 3: Item Field Mapping Reference

| [`YahooAuctionItem`](file:///c:/Users/AR/Desktop/mercari-jp-mcp/temp_download/yahoo_auctions/models.py) Attribute | Raw JSON Field in `raw_items[i]` | Transformation / Logic |
| :--- | :--- | :--- |
| `id` | `item['auctionId']` | Direct string |
| `title` | `item['title']` | Stripped string |
| `price` | `item['price']` or `item['currentPrice']` | Cast to `int` (defaults to 0) |
| `buy_now_price` | `item['buyNowPrice']` | Cast to `int` if not `None`, else `None` |
| `bid_count` | `item['bidCount']` | Cast to `int` (defaults to 0) |
| `end_time` | `item['endTime']` | ISO-8601 string (e.g. `"2026-09-30T20:02:04+09:00"`) |
| `end_timestamp` | Derived from `item['endTime']` or `item['etc']` | `int(datetime.fromisoformat(end_time).timestamp())` or regex `etm=(\d+)` |
| `time_remaining` | `item['leftTime']` (in seconds) | Formatted via `format_time_remaining()` into `Xd Yh` / `Xh Ym` / `Xm` / `Ended` |
| `seller_id` | `item['seller']['userId']` | Direct string |
| `seller_name` | `item['seller']['displayName']` | Direct string |
| `thumbnail` | `item['imageUrl']` | Direct Yahoo CDN image URL |
| `url` | Derived | If `item.get('isFleamarketItem')` &rarr; `https://paypayfleamarket.yahoo.co.jp/item/{id}`<br>Else &rarr; `https://page.auctions.yahoo.co.jp/auction/{id}` |
| `free_shipping` | `item['isFreeShipping']` | Cast to `bool` |
| `unused` | `item['itemCondition']` | `'UNUSED' in cond or cond == 'NEW'` |
| `is_new` | `item['isNewArrival']` or `item['itemCondition']` | `bool(item.get('isNewArrival')) or cond == 'NEW'` |
| `status` | `item['endTimeStatus']`, `isClosed`, or `status` arg | `'sold'` if `status == 'sold'` or `endTimeStatus in ('CLOSED', 'ENDED', 'TODAY')` or `isClosed` else `'live'` |

---

## 5. Item Detail Parsing Specification

Located in [`temp_download/yahoo_auctions/item.py`](file:///c:/Users/AR/Desktop/mercari-jp-mcp/temp_download/yahoo_auctions/item.py).

### Dual-URL Resolution
Because listings can belong to Yahoo Auctions or PayPay Fleamarket:
1. If `item_id.startswith('z')`: Try `https://paypayfleamarket.yahoo.co.jp/item/{id}` first, then `https://page.auctions.yahoo.co.jp/auction/{id}`.
2. Otherwise: Try `https://page.auctions.yahoo.co.jp/auction/{id}` first, then `https://paypayfleamarket.yahoo.co.jp/item/{id}`.

### Multi-Tier Extraction Hierarchy

```
[Tier 1: __NEXT_DATA__ JSON]
 ├── Yahoo Auction: init['item']['detail']['item']
 │    ├── title, price, bidorbuy, bids, formattedEndTime, leftTime, img[]
 └── PayPay Fleamarket: init['itemsState']['items']['item']
      ├── title, price, description, status, seller, images[]
      │
[Tier 2: Schema.org JSON-LD (<script type="application/ld+json">)]
 ├── @type == "Product"
 └── name (title), image[], offers.price, description, offers.itemCondition
      │
[Tier 3: OpenGraph & HTML Meta Tags]
 └── <meta property="og:title">, <meta property="og:image">, <meta property="og:description">
```

---

## 6. Anti-Bot, Headers & Throttling Rules

Located in [`temp_download/yahoo_auctions/client.py`](file:///c:/Users/AR/Desktop/mercari-jp-mcp/temp_download/yahoo_auctions/client.py).

1. **Rate Limiting:**
   - Yahoo will block rapid concurrent hits or soft-404 requests.
   - `MIN_DELAY = 1.5` seconds enforced via `REQUEST_LOCK = threading.Lock()`.
2. **Session Warmup:**
   - `_warmup()` makes an initial request to `https://auctions.yahoo.co.jp/` to establish cookies before sending search requests.
3. **Soft-404 Recognition:**
   - Yahoo sometimes responds with HTTP 404 or 403 on anti-bot checks while still providing page data, or on valid empty search results ("見つかりません").
   - `_RESULTS_MARKERS = ('__NEXT_DATA__', 'Product__imageLink', 'auctionId', 'itemsState')` ensures that valid Next.js result pages are parsed and not discarded as errors.
4. **Retry & Backoff:**
   - `MAX_RETRIES = 3`. On retry, the session is reset via `_reset_session()` and cookies are cleared.

---

## 7. How to Test & Debug (Quick Reference for AI Agents)

Whenever making changes, test with this quick script in Python:

```bash
uv run python -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'temp_download')
from yahoo_auctions import search_yahoo, get_item_detail

# 1. Test Live Search
live = search_yahoo('switch', status='live', limit=3)
print('Live search count:', len(live))
assert len(live) > 0 and live[0].title and live[0].thumbnail

# 2. Test Sold Search
sold = search_yahoo('switch', status='sold', limit=3)
print('Sold search count:', len(sold))
assert len(sold) > 0 and sold[0].title and sold[0].thumbnail

# 3. Test Detail
detail = get_item_detail(live[0].id)
print('Live detail title:', detail.get('title'), 'Images:', len(detail.get('images', [])))
assert detail.get('title') and len(detail.get('images', [])) > 0
print('ALL CHECKS PASSED!')
"
```

If Yahoo changes their page structure in the future:
1. Run `client.fetch(url)` on a search URL and save to a file `debug_search.html`.
2. Check if `<script id="__NEXT_DATA__">` still exists.
3. If it exists, inspect the JSON keys with:
   `python -c "import json, re; m=re.search(r'<script id=\"__NEXT_DATA__\"[^>]*>(.*?)</script>', open('debug_search.html', encoding='utf-8').read(), re.DOTALL); d=json.loads(m.group(1)); print(list(d['props']['pageProps']['initialState']['search']['items'].keys()))"`
4. Adjust the key path in `temp_download/yahoo_auctions/search.py` accordingly.
