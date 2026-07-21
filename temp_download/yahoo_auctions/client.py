import time
import requests

SESSION = None
LAST_REQUEST = 0
MIN_DELAY = 1.5

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Cache-Control': 'max-age=0',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
}


def get_session():
    global SESSION
    if SESSION is None:
        SESSION = requests.Session()
        SESSION.headers.update(HEADERS)
        _warmup()
    return SESSION


def _warmup():
    try:
        SESSION.get('https://auctions.yahoo.co.jp/', timeout=10, allow_redirects=True)
        time.sleep(1)
    except Exception:
        pass


def fetch(url, timeout=15):
    global LAST_REQUEST
    session = get_session()
    elapsed = time.time() - LAST_REQUEST
    if elapsed < MIN_DELAY:
        time.sleep(MIN_DELAY - elapsed)
    LAST_REQUEST = time.time()
    resp = session.get(url, timeout=timeout)
    resp.raise_for_status()
    resp.encoding = 'utf-8'
    return resp.text
