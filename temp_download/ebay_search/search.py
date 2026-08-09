import re
import requests as http_requests
import base64
from .models import EbayItem

LEGACY_ID_RE = re.compile(r'(?:/itm/|item=)(\d+)', re.I)

OAUTH_URL = 'https://api.ebay.com/identity/v1/oauth2/token'
BROWSE_API_URL = 'https://api.ebay.com/buy/browse/v1/item_summary/search'
ITEM_API_URL = 'https://api.ebay.com/buy/browse/v1/item'

CONDITION_MAP = {
    'new': 'NEW',
    'used': 'USED',
    'open_box': 'OPEN_BOX',
    'refurbished': 'REFURBISHED',
    'for_parts': 'FOR_PARTS',
}

CONDITION_MAP = {
    1000: 'New',
    1500: 'New with Tags',
    1750: 'New with Box',
    2000: 'Manufacturer Refurbished',
    2500: 'Seller Refurbished',
    2750: 'Like New',
    3000: 'Used',
    4000: 'Very Good',
    5000: 'Good',
    6000: 'Acceptable',
    7000: 'For Parts or Not Working',
    8000: 'Brand New',
}

SORT_MAP = {
    'best_match': '',
    'price_asc': 'price',
    'price_desc': '-price',
    'distance': 'distance',
}


def _get_oauth_token(app_id: str, cert_id: str) -> str:
    credentials = f'{app_id}:{cert_id}'
    encoded = base64.b64encode(credentials.encode('utf-8')).decode('utf-8')

    headers = {
        'Authorization': f'Basic {encoded}',
        'Content-Type': 'application/x-www-form-urlencoded',
    }
    data = {
        'grant_type': 'client_credentials',
        'scope': 'https://api.ebay.com/oauth/api_scope',
    }

    resp = http_requests.post(OAUTH_URL, headers=headers, data=data, timeout=15)
    resp.raise_for_status()
    token_data = resp.json()
    return token_data['access_token']


def extract_legacy_item_id(url_or_id: str) -> str:
    s = (url_or_id or '').strip()
    if not s:
        return ''
    if s.isdigit():
        return s
    match = LEGACY_ID_RE.search(s)
    if match:
        return match.group(1)
    if '|' in s:
        parts = s.split('|')
        if len(parts) >= 2 and parts[1].isdigit():
            return parts[1]
    return ''


def _parse_single_item(item: dict):
    item_id = item.get('itemId', '')
    if not item_id:
        return None

    title = item.get('title', 'No Title')

    price_obj = item.get('price', {}) or {}
    value = price_obj.get('value', '0')
    price = float(value) if value else 0
    currency = price_obj.get('currency', 'USD')

    thumbnail = ''
    images = item.get('thumbnailImages', []) or item.get('image', {}).get('imageUrl')
    if isinstance(images, list) and images:
        thumbnail = images[0].get('imageUrl', '')
    elif isinstance(images, str):
        thumbnail = images
    if not thumbnail:
        image_obj = item.get('image', {})
        if isinstance(image_obj, dict):
            thumbnail = image_obj.get('imageUrl', '')

    url = item.get('itemWebUrl', '') or item.get('itemAffiliateWebUrl', '')

    condition = item.get('condition', '')
    condition_id = item.get('conditionId', None)
    if not condition and condition_id:
        try:
            condition = CONDITION_MAP.get(int(condition_id), f'Condition {condition_id}')
        except (TypeError, ValueError):
            pass

    if condition == 'New \u2013 Open box':
        condition = 'Open Box (New)'
    if condition in ('Ungraded', 'Unspecified', ''):
        t = title.lower()
        electronics = ['laptop', 'phone', 'pc', 'computer', 'iphone', 'ipad', 'camera',
                       'watch', 'console', 'ps4', 'ps5', 'switch', 'tv', 'monitor',
                       'keyboard', 'drone', 'gpu', 'cpu', 'macbook', 'tablet', 'printer',
                       'console', 'xbox', 'airpods', 'speaker', 'headphone', 'hard drive']
        junk_terms = ['\u30b8\u30e3\u30f3\u30af', 'broken', 'for parts', 'parts only',
                      '\u8a33\u3042\u308a', '\u4e0d\u826f', 'not working', 'untested',
                      'damaged', 'dead']
        if any(m in t for m in junk_terms) or ('junk' in t and any(e in t for e in electronics)):
            condition = 'For parts or not working'
        elif condition == 'Ungraded':
            condition = 'Used'
        elif condition == 'Unspecified':
            condition = 'Not Specified'

    listing_type = 'FIXED_PRICE'
    is_auction = False
    bid_count = item.get('bidCount')
    buying_options = item.get('buyingOptions', ['FIXED_PRICE'])
    if 'AUCTION' in buying_options:
        listing_type = 'AUCTION'
        is_auction = True

    bin_price = None
    if 'FIXED_PRICE' in buying_options:
        bin_price_obj = item.get('price', {})
        bin_val = bin_price_obj.get('value', None)
        bin_price = float(bin_val) if bin_val else None
    current_bid = item.get('currentBidPrice', {}) or {}
    if is_auction and current_bid.get('value'):
        price = float(current_bid.get('value'))

    seller = item.get('seller', {})
    seller_name = seller.get('username', '')
    feedback = seller.get('feedbackPercentage', '')

    shipping = item.get('shippingOptions', [{}])[0] if item.get('shippingOptions') else {}
    shipping_type = shipping.get('shippingCostType', '')
    shipping_cost = None
    free_shipping = False
    if shipping_type == 'CALCULATED':
        shipping_cost = None
    else:
        cost_val = shipping.get('shippingCost', {}).get('value', None)
        if cost_val is not None:
            cost_float = float(cost_val)
            if cost_float == 0:
                free_shipping = True
            else:
                shipping_cost = cost_float

    location = item.get('itemLocation', {}).get('city', '')
    if not location:
        location = item.get('itemLocation', {}).get('country', '')

    top_rated = item.get('topRatedBuyingExperience', False)

    category_name = ''
    categories = item.get('categories', [])
    if categories:
        category_name = categories[0].get('categoryName', '')

    return EbayItem(
        id=item_id,
        title=title,
        price=price,
        currency=currency,
        condition=condition,
        condition_id=condition_id,
        thumbnail=thumbnail,
        url=url,
        listing_type=listing_type,
        bid_count=bid_count,
        buy_it_now_price=bin_price,
        seller_name=seller_name,
        seller_rating=str(feedback) if feedback else '',
        location=location,
        shipping_cost=shipping_cost,
        shipping_type=shipping_type,
        free_shipping=free_shipping,
        is_auction=is_auction,
        top_rated_seller=top_rated,
        category_name=category_name,
    )


def parse_browse_results(data: dict) -> list:
    items = []
    for item in data.get('itemSummaries', []):
        try:
            parsed = _parse_single_item(item)
            if parsed:
                items.append(parsed)
        except Exception:
            continue
    return items


def get_ebay_item(
    app_id: str,
    cert_id: str,
    url_or_id: str,
    global_id: str = 'EBAY-US',
    token: str = None,
) -> EbayItem:
    legacy_id = extract_legacy_item_id(url_or_id)
    if not legacy_id:
        raise Exception('Invalid eBay item URL or ID')

    if not cert_id:
        import os
        cert_id = os.environ.get('EBAY_CERT_ID', '')
    if not cert_id:
        raise Exception('eBay Cert ID is required')

    if not token:
        token = _get_oauth_token(app_id, cert_id)

    headers = {
        'Authorization': f'Bearer {token}',
        'X-EBAY-C-MARKETPLACE-ID': global_id,
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    }

    resp = http_requests.get(
        f'{ITEM_API_URL}/get_item_by_legacy_id',
        headers=headers,
        params={'legacy_item_id': legacy_id},
        timeout=30,
    )
    if resp.status_code == 404:
        resp = http_requests.get(
            f'{ITEM_API_URL}/v1|{legacy_id}|0',
            headers=headers,
            timeout=30,
        )
    if resp.status_code == 403:
        raise Exception('eBay API returned 403. Check Browse API access.')
    if resp.status_code == 401:
        raise Exception('eBay API returned 401 Unauthorized. Check App ID and Cert ID.')
    if resp.status_code == 404:
        raise Exception('eBay item not found')
    resp.raise_for_status()

    parsed = _parse_single_item(resp.json())
    if not parsed:
        raise Exception('Could not parse eBay item details')
    return parsed


def search_ebay(
    app_id: str,
    keyword: str,
    min_price: float = None,
    max_price: float = None,
    condition: str = '',
    sort: str = 'best_match',
    limit: int = 20,
    page: int = 1,
    global_id: str = 'EBAY-US',
    category_id: str = '',
    bin_only: bool = False,
    cert_id: str = None,
    item_location: str = '',
) -> list:
    if not cert_id:
        import os
        cert_id = os.environ.get('EBAY_CERT_ID', '')

    if not cert_id:
        raise Exception('eBay Cert ID (Client Secret) is required for OAuth. Set EBAY_CERT_ID env var or pass cert_id parameter.')

    token = _get_oauth_token(app_id, cert_id)

    headers = {
        'Authorization': f'Bearer {token}',
        'X-EBAY-C-MARKETPLACE-ID': global_id,
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    }

    params = {
        'q': keyword,
        'limit': str(min(limit, 200)),
        'offset': str((page - 1) * limit),
    }

    sort_val = SORT_MAP.get(sort)
    if sort_val:
        params['sort'] = sort_val

    if category_id:
        params['category_ids'] = category_id

    filters = []
    if condition:
        cond_val = CONDITION_MAP.get(condition)
        if cond_val:
            filters.append(f'conditions:{cond_val}')

    if min_price is not None or max_price is not None:
        price_filter = 'price:['
        price_filter += str(min_price) if min_price is not None else ''
        price_filter += '..'
        price_filter += str(max_price) if max_price is not None else ''
        price_filter += ']'
        filters.append(price_filter)

    if bin_only:
        filters.append('buyingOptions:{FIXED_PRICE}')

    if item_location:
        filters.append(f'itemLocationCountry:{item_location}')

    if filters:
        params['filter'] = ','.join(filters)

    resp = http_requests.get(
        BROWSE_API_URL,
        headers=headers,
        params=params,
        timeout=30,
    )

    if resp.status_code == 403:
        raise Exception('eBay API returned 403. Check that your OAuth scope has access to the Browse API.')
    if resp.status_code == 401:
        raise Exception('eBay API returned 401 Unauthorized. Check your App ID and Cert ID.')

    resp.raise_for_status()
    data = resp.json()

    return parse_browse_results(data)
