import re
import json
from .client import fetch

DETAIL_URL = 'https://page.auctions.yahoo.co.jp/auction/{item_id}'
FLEAMARKET_URL = 'https://paypayfleamarket.yahoo.co.jp/item/{item_id}'


def _extract_next_data(html: str) -> dict:
    """Extract the __NEXT_DATA__ JSON embedded in the page."""
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.DOTALL)
    if not m:
        return {}
    try:
        return json.loads(m.group(1))
    except (json.JSONDecodeError, KeyError):
        return {}


def _extract_json_ld(html: str) -> list:
    """Extract JSON-LD structured data."""
    results = []
    for m in re.finditer(
        r'<script type="application/ld\+json"[^>]*>(.*?)</script>', html, re.DOTALL
    ):
        try:
            data = json.loads(m.group(1))
            if isinstance(data, dict):
                results.append(data)
            elif isinstance(data, list):
                results.extend(data)
        except json.JSONDecodeError:
            continue
    return results


def get_item_detail(item_id: str) -> dict:
    result = {
        'id': item_id,
        'title': '',
        'description': '',
        'price': 0,
        'buy_now_price': None,
        'condition': '',
        'seller_name': '',
        'seller_id': '',
        'seller_rating': '',
        'seller_reviews': 0,
        'end_time': '',
        'end_timestamp': None,
        'left_time': '',
        'bid_count': 0,
        'bidders': 0,
        'watchlist': 0,
        'status': '',
        'images': [],
        'shipping': '',
    }

    # Decide order of URLs to fetch based on item_id pattern
    # Fleamarket items commonly start with 'z'
    urls_to_try = []
    if item_id.startswith('z'):
        urls_to_try = [
            FLEAMARKET_URL.format(item_id=item_id),
            DETAIL_URL.format(item_id=item_id)
        ]
    else:
        urls_to_try = [
            DETAIL_URL.format(item_id=item_id),
            FLEAMARKET_URL.format(item_id=item_id)
        ]

    html = ''
    for url in urls_to_try:
        try:
            h = fetch(url)
            if h and len(h) > 500:
                html = h
                # If page contains explicit not found / 404 in NEXT_DATA, try fallback url
                if '"error":{"status":404' in h:
                    continue
                break
        except Exception:
            continue

    if not html:
        return result

    # 1. Try __NEXT_DATA__
    nd = _extract_next_data(html)
    if nd:
        try:
            props = nd.get('props', {})
            page_props = props.get('pageProps', {})
            init = page_props.get('initialState') or props.get('initialState', {})

            # A. Yahoo Auctions detail structure
            auction_item = init.get('item', {}).get('detail', {}).get('item', {})
            if auction_item and isinstance(auction_item, dict):
                result['title'] = auction_item.get('title', '')
                result['price'] = auction_item.get('price', 0)
                result['buy_now_price'] = auction_item.get('bidorbuy', None)
                result['bid_count'] = auction_item.get('bids', 0)
                result['bidders'] = auction_item.get('biddersNum', 0)
                result['end_time'] = auction_item.get('formattedEndTime', '')
                result['end_timestamp'] = auction_item.get('endTime', '')
                lt = auction_item.get('leftTime', 0)
                if lt and lt > 0:
                    days = int(lt // 86400)
                    hours = int((lt % 86400) // 3600)
                    parts = []
                    if days:
                        parts.append(f'{days}d')
                    if hours:
                        parts.append(f'{hours}h')
                    result['left_time'] = ' '.join(parts)
                elif lt and lt <= 0:
                    result['left_time'] = 'Ended'
                result['status'] = auction_item.get('status', '')
                result['watchlist'] = auction_item.get('watchListNum', 0)
                result['shipping'] = auction_item.get('chargeForShipping', '')

                seller = auction_item.get('seller', {})
                if seller:
                    result['seller_name'] = seller.get('displayName', '')
                    result['seller_id'] = seller.get('aucUserId', '')
                    rating = seller.get('rating', {})
                    result['seller_rating'] = rating.get('goodRating', '')
                    result['seller_reviews'] = rating.get('summary', 0)

                images = auction_item.get('img', [])
                if isinstance(images, list):
                    result['images'] = [
                        img.get('image', '')
                        for img in images
                        if isinstance(img, dict) and img.get('image')
                    ]

            # B. PayPay Fleamarket structure
            fleamarket_item = init.get('itemsState', {}).get('items', {}).get('item', {})
            if fleamarket_item and isinstance(fleamarket_item, dict):
                if not result['title']:
                    result['title'] = fleamarket_item.get('title', '')
                if not result['price']:
                    result['price'] = fleamarket_item.get('price', 0)
                    result['buy_now_price'] = fleamarket_item.get('price', 0)
                if not result['status']:
                    result['status'] = fleamarket_item.get('status', '')
                if not result['description']:
                    result['description'] = fleamarket_item.get('description', '')
                seller = fleamarket_item.get('seller', {})
                if seller and not result['seller_name']:
                    result['seller_name'] = seller.get('name', '')
                    result['seller_id'] = seller.get('id', '')
                fimages = fleamarket_item.get('images', [])
                if fimages and not result['images']:
                    result['images'] = [
                        img.get('url', '') if isinstance(img, dict) else str(img)
                        for img in fimages
                        if img
                    ]
        except (KeyError, TypeError):
            pass

    # 2. JSON-LD fallback for description, title, images, price, and condition
    for ld in _extract_json_ld(html):
        if ld.get('@type') == 'Product':
            if not result['title'] and ld.get('name'):
                result['title'] = ld['name']
            if not result['description'] and ld.get('description'):
                result['description'] = ld['description']
            if not result['images'] and ld.get('image'):
                imgs = ld['image']
                result['images'] = imgs if isinstance(imgs, list) else [imgs]
            offers = ld.get('offers', {})
            if isinstance(offers, dict):
                if not result['price'] and offers.get('price'):
                    try:
                        result['price'] = int(float(offers['price']))
                    except (ValueError, TypeError):
                        pass
                if not result['condition']:
                    cond_url = offers.get('itemCondition', '')
                    cond_map = {
                        'newcondition': 'New',
                        'usedcondition': 'Used',
                        'refurbishedcondition': 'Refurbished',
                        'damagedcondition': 'Damaged',
                    }
                    for key, val in cond_map.items():
                        if key in cond_url.lower():
                            result['condition'] = val
                            break

    # 3. HTML meta tag fallback
    if not result['title']:
        m = re.search(r'<meta property="og:title" content="([^"]+)"', html)
        if m:
            t = m.group(1)
            t = re.sub(r'\s*-\s*Yahoo!オークション.*$', '', t)
            result['title'] = t.strip()
        else:
            m2 = re.search(r'<title>(.*?)</title>', html)
            if m2:
                t = m2.group(1)
                t = re.sub(r'Yahoo!オークション\s*-\s*', '', t)
                t = re.sub(r'\s*-\s*Yahoo!オークション.*$', '', t)
                result['title'] = t.strip()

    if not result['images']:
        m = re.search(r'<meta property="og:image" content="([^"]+)"', html)
        if m:
            result['images'] = [m.group(1)]

    if not result['description']:
        m = re.search(r'<meta property="og:description" content="([^"]+)"', html)
        if m:
            result['description'] = m.group(1)

    return result
