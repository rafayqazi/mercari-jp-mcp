import re
from urllib.parse import urlencode, urlparse, parse_qs, urljoin
from bs4 import BeautifulSoup
from .client import fetch
from .models import YahooAuctionItem

BASE_URL = 'https://auctions.yahoo.co.jp/search/search'


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

    # auccat: 'new' for live listings, 'closed' for sold
    if status == 'sold':
        params['auccat'] = 'closed'
    else:
        params['auccat'] = 'new'

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

    return BASE_URL + '?' + urlencode(params)


def parse_search_results(html: str) -> list:
    soup = BeautifulSoup(html, 'lxml')
    items = []

    product_cards = soup.select('li.Product')
    if not product_cards:
        return items

    for card in product_cards:
        try:
            item = _parse_card(card)
            if item:
                items.append(item)
        except Exception:
            continue

    return items


def _parse_card(card) -> YahooAuctionItem:
    image_link = card.select_one('a.Product__imageLink')
    if not image_link:
        return None

    item_id = image_link.get('data-auction-id', '')
    if not item_id:
        return None

    url = image_link.get('href', '')
    if not url:
        url = f'https://page.auctions.yahoo.co.jp/auction/{item_id}'

    thumb_img = card.select_one('img.Product__imageData')
    thumbnail = thumb_img.get('src', '') if thumb_img else ''

    free_shipping = bool(card.select_one('.Product__icon--freeShipping'))
    unused = bool(card.select_one('.Product__icon--unused'))
    is_new = bool(card.select_one('.Product__icon--new'))

    title_elem = card.select_one('h3.Product__title a.Product__titleLink')
    title = title_elem.get('title', '') or title_elem.get_text(strip=True) if title_elem else ''

    price = 0
    buy_now_price = None

    price_values = card.select('.Product__priceValue')
    price_labels = card.select('.Product__label')

    # Parse price info
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

    # Determine status
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
        keyword=keyword,
        min_price=min_price,
        max_price=max_price,
        status=status,
        condition=condition,
        sort=sort,
        page=page,
        limit=limit,
    )
    html = fetch(url)
    return parse_search_results(html)
