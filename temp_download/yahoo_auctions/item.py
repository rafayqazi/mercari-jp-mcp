import re
import json
from .client import fetch

DETAIL_URL = 'https://page.auctions.yahoo.co.jp/auction/{item_id}'


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
    url = DETAIL_URL.format(item_id=item_id)
    html = fetch(url)

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

    # Try __NEXT_DATA__ first (rich structured data)
    nd = _extract_next_data(html)
    if nd:
        try:
            item = (
                nd.get('props', {})
                .get('pageProps', {})
                .get('initialState', {})
                .get('item', {})
                .get('detail', {})
                .get('item', {})
            )
            if item:
                result['title'] = item.get('title', '')
                result['price'] = item.get('price', 0)
                result['buy_now_price'] = item.get('bidorbuy', None)
                result['bid_count'] = item.get('bids', 0)
                result['bidders'] = item.get('biddersNum', 0)
                result['end_time'] = item.get('formattedEndTime', '')
                result['end_timestamp'] = item.get('endTime', '')
                result['left_time'] = ''
                lt = item.get('leftTime', 0)
                if lt:
                    days = int(lt // 86400)
                    hours = int((lt % 86400) // 3600)
                    parts = []
                    if days:
                        parts.append(f'{days}d')
                    if hours:
                        parts.append(f'{hours}h')
                    result['left_time'] = ' '.join(parts)
                result['status'] = item.get('status', '')
                result['watchlist'] = item.get('watchListNum', 0)
                result['shipping'] = item.get('chargeForShipping', '')

                seller = item.get('seller', {})
                if seller:
                    result['seller_name'] = seller.get('displayName', '')
                    result['seller_id'] = seller.get('aucUserId', '')
                    rating = seller.get('rating', {})
                    result['seller_rating'] = rating.get('goodRating', '')
                    result['seller_reviews'] = rating.get('summary', 0)

                images = item.get('img', [])
                result['images'] = [img.get('image', '') for img in images if img.get('image')]

                # Description is typically in the JSON-LD rather than __NEXT_DATA__
        except (KeyError, TypeError) as e:
            pass

    # Try JSON-LD for description (__NEXT_DATA__ may not have it)
    if not result['description']:
        for ld in _extract_json_ld(html):
            if ld.get('@type') == 'Product':
                result['description'] = ld.get('description', '')
                if not result['condition']:
                    cond_url = ld.get('offers', {}).get('itemCondition', '')
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

    return result
