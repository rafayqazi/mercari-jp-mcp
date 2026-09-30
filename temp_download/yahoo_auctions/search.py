import re
import json
from datetime import datetime
from urllib.parse import urlencode
from bs4 import BeautifulSoup
from .client import fetch
from .models import YahooAuctionItem

BASE_URL = 'https://auctions.yahoo.co.jp/search/search'
CLOSED_URL = 'https://auctions.yahoo.co.jp/closedsearch/closedsearch'


def clean_search_keyword(keyword: str) -> str:
    """Trim clutter so Yahoo matches the meaningful product name.

    Yahoo search OR-matches tokens, so long item titles (e.g.
    "JOJOVELLER (Complete Limited Edition) (Used)") pull in unrelated
    listings matching just "Limited Edition" / "Used". When the keyword
    contains parenthetical clauses or size separators, strip them, drop
    pure-number tokens (sizes like 1/7), and keep only the first two
    meaningful words. Clean short queries pass through unchanged.
    """
    if not keyword:
        return keyword
    if not re.search(r'[()/:;,]', keyword):
        return keyword
    k = re.sub(r'\([^)]*\)', ' ', keyword)
    k = re.sub(r'[/\\]', ' ', k)
    k = re.sub(r'[:;,."\'*~!?#]', ' ', k)
    tokens = [t for t in k.split() if t]
    return ' '.join(tokens[:2])


def build_search_url(
    keyword: str,
    min_price: int = None,
    max_price: int = None,
    status: str = 'live',
    condition: str = '',
    sort: str = 'new',
    page: int = 1,
    limit: int = 20
) -> str:
    params = {
        'p': keyword,
        'ei': 'UTF-8',
        'b': str((page - 1) * limit + 1),
        'n': str(limit),
    }

    if min_price is not None:
        params['min'] = str(min_price)
    if max_price is not None:
        params['max'] = str(max_price)

    if condition:
        params['istatus'] = condition

    sort_map = {
        'new': '-new',      # newest
        'end': 'end',        # ending soon
        'price': 'price',    # price asc
        '-price': '-price',  # price desc
        'bid': '-bid',       # most bids
    }
    if sort in sort_map:
        params['s'] = sort_map[sort]

    if status == 'sold':
        return CLOSED_URL + '?' + urlencode(params)
    else:
        params['auccat'] = 'new'
        return BASE_URL + '?' + urlencode(params)


def format_time_remaining(left_time) -> str:
    """Format leftTime (seconds) into a compact, human-readable string."""
    if left_time is None:
        return ''
    try:
        lt = float(left_time)
        if lt <= 0:
            return 'Ended'
        if lt < 60:
            return f'{int(lt)}s'
        if lt < 3600:
            return f'{int(lt // 60)}m'
        if lt < 86400:
            h = int(lt // 3600)
            m = int((lt % 3600) // 60)
            return f'{h}h {m}m' if m else f'{h}h'
        d = int(lt // 86400)
        h = int((lt % 86400) // 3600)
        return f'{d}d {h}h' if h else f'{d}d'
    except (ValueError, TypeError):
        return ''


def _parse_next_data_items(html: str) -> list:
    """Parse search listings from embedded __NEXT_DATA__ JSON in modern Yahoo Auctions."""
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.DOTALL)
    if not m:
        return []

    try:
        nd = json.loads(m.group(1))
    except (json.JSONDecodeError, KeyError):
        return []

    props = nd.get('props', {})
    page_props = props.get('pageProps', {})
    init = page_props.get('initialState') or props.get('initialState', {})
    search = init.get('search', {})
    items_state = search.get('items', {})

    raw_items = items_state.get('listing', {}).get('items', [])
    if not raw_items:
        raw_items = items_state.get('featuredListing', {}).get('items', [])

    results = []
    for item in raw_items:
        item_id = str(item.get('auctionId') or '').strip()
        if not item_id:
            continue

        title = str(item.get('title') or '').strip()
        price = item.get('price')
        if price is None:
            price = item.get('currentPrice', 0)
        try:
            price = int(price)
        except (ValueError, TypeError):
            price = 0

        buy_now = item.get('buyNowPrice')
        try:
            buy_now_price = int(buy_now) if buy_now is not None else None
        except (ValueError, TypeError):
            buy_now_price = None

        try:
            bid_count = int(item.get('bidCount') or 0)
        except (ValueError, TypeError):
            bid_count = 0

        end_time = str(item.get('endTime') or '')
        end_timestamp = None
        if end_time:
            try:
                end_timestamp = int(datetime.fromisoformat(end_time).timestamp())
            except Exception:
                pass
        if not end_timestamp:
            etm_m = re.search(r'etm=(\d+)', str(item.get('etc') or ''))
            if etm_m:
                try:
                    end_timestamp = int(etm_m.group(1))
                except ValueError:
                    pass

        time_remaining = format_time_remaining(item.get('leftTime'))

        seller = item.get('seller') or {}
        seller_id = str(seller.get('userId') or '')
        seller_name = str(seller.get('displayName') or '')

        thumbnail = str(item.get('imageUrl') or '').strip()

        is_fleamarket = bool(item.get('isFleamarketItem', False))
        if is_fleamarket:
            url = f'https://paypayfleamarket.yahoo.co.jp/item/{item_id}'
        else:
            url = f'https://page.auctions.yahoo.co.jp/auction/{item_id}'

        free_shipping = bool(item.get('isFreeShipping', False))
        cond = str(item.get('itemCondition') or '').upper()
        unused = 'UNUSED' in cond or cond == 'NEW'
        is_new = bool(item.get('isNewArrival', False)) or cond == 'NEW'

        status = 'live'
        if item.get('endTimeStatus') in ('CLOSED', 'ENDED', 'TODAY') or item.get('isClosed', False) or '落札' in title:
            status = 'sold'

        results.append(YahooAuctionItem(
            id=item_id,
            title=title,
            price=price,
            buy_now_price=buy_now_price,
            bid_count=bid_count,
            end_time=end_time,
            end_timestamp=end_timestamp,
            time_remaining=time_remaining,
            seller_id=seller_id,
            seller_name=seller_name,
            thumbnail=thumbnail,
            url=url,
            free_shipping=free_shipping,
            unused=unused,
            is_new=is_new,
            status=status,
        ))

    return results


def _parse_card(card) -> YahooAuctionItem:
    """Legacy HTML card parser fallback."""
    image_link = card.select_one('a.Product__imageLink') or card.select_one('a[href*="/auction/"]')
    if not image_link:
        return None

    item_id = image_link.get('data-auction-id', '')
    url = image_link.get('href', '')
    if not item_id and url:
        m = re.search(r'/auction/([a-zA-Z0-9]+)', url)
        if m:
            item_id = m.group(1)

    if not item_id:
        return None

    if not url:
        url = f'https://page.auctions.yahoo.co.jp/auction/{item_id}'

    thumb_img = card.select_one('img.Product__imageData') or card.select_one('img')
    thumbnail = thumb_img.get('src', '') if thumb_img else ''

    free_shipping = bool(card.select_one('.Product__icon--freeShipping'))
    unused = bool(card.select_one('.Product__icon--unused'))
    is_new = bool(card.select_one('.Product__icon--new'))

    title_elem = card.select_one('h3.Product__title a.Product__titleLink') or card.select_one('h3 a') or card.select_one('a[title]')
    title = title_elem.get('title', '') or title_elem.get_text(strip=True) if title_elem else ''

    price = 0
    buy_now_price = None

    price_values = card.select('.Product__priceValue')
    price_labels = card.select('.Product__label')

    for label, val in zip(price_labels, price_values):
        label_text = label.get_text(strip=True)
        val_text = val.get_text(strip=True)
        val_clean = re.sub(r'[^\d]', '', val_text)
        val_int = int(val_clean) if val_clean else 0
        if '現在' in label_text:
            price = val_int
        elif '即決' in label_text:
            buy_now_price = val_int

    bid_count = 0
    bid_elem = card.select_one('dd.Product__bid')
    if bid_elem:
        bid_text = bid_elem.get_text(strip=True)
        bid_clean = re.sub(r'[^\d]', '', bid_text)
        bid_count = int(bid_clean) if bid_clean else 0

    time_elem = card.select_one('dd.Product__time')
    time_remaining = time_elem.get_text(strip=True) if time_elem else ''

    bonus = card.select_one('.Product__bonus')
    seller_id = bonus.get('data-auction-auc-seller-id', '') if bonus else ''
    end_timestamp = None
    if bonus:
        end_ts = bonus.get('data-auction-endtime', '')
        if end_ts:
            try:
                end_timestamp = int(end_ts)
            except ValueError:
                pass

    status = 'live'
    if card.select_one('.Product__icon--closed') or '落札' in title:
        status = 'sold'

    return YahooAuctionItem(
        id=item_id,
        title=title,
        price=price,
        buy_now_price=buy_now_price,
        bid_count=bid_count,
        end_timestamp=end_timestamp,
        time_remaining=time_remaining,
        seller_id=seller_id,
        thumbnail=thumbnail,
        url=url,
        free_shipping=free_shipping,
        unused=unused,
        is_new=is_new,
        status=status,
    )


def parse_search_results(html: str) -> list:
    """Parse search results from Yahoo Auctions, supporting both modern Next.js and legacy HTML."""
    if not html:
        return []

    # 1. Try modern Next.js structured data first
    items = _parse_next_data_items(html)
    if items:
        return items

    # 2. Legacy HTML DOM fallback
    soup = BeautifulSoup(html, 'lxml')
    items = []
    product_cards = soup.select('li.Product')
    for card in product_cards:
        try:
            item = _parse_card(card)
            if item:
                items.append(item)
        except Exception:
            continue

    return items


def search_yahoo(
    keyword: str,
    min_price: int = None,
    max_price: int = None,
    status: str = 'live',
    condition: str = '',
    sort: str = 'new',
    page: int = 1,
    limit: int = 20,
) -> list:
    url = build_search_url(
        keyword=clean_search_keyword(keyword),
        min_price=min_price,
        max_price=max_price,
        status=status,
        condition=condition,
        sort=sort,
        page=page,
        limit=limit,
    )
    html = fetch(url)
    items = parse_search_results(html)
    if limit and len(items) > limit:
        return items[:limit]
    return items
