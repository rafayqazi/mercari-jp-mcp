import time
import threading
import requests

SESSION = None
LAST_REQUEST = 0
MIN_DELAY = 1.5
REQUEST_LOCK = threading.Lock()
MAX_RETRIES = 3

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

# Yahoo answers search requests with:
#  - HTTP 200 + real results page (normal case)
#  - HTTP 404 + real results page (soft-404 anti-bot, still valid results)
#  - HTTP 404 + "no results" page (見つかりません = keyword has no listings)
# Treat the first two as success and the last as a valid empty result.
_RESULTS_MARKER = 'Product__imageLink'
_NO_RESULTS_MARKERS = ('見つかりません', '検索結果がありません')


def get_session():
    global SESSION
    if SESSION is None:
        SESSION = requests.Session()
        SESSION.headers.update(HEADERS)
        _warmup()
    return SESSION


def _reset_session():
    """Drop the session so the next request gets a fresh cookie jar + warmup."""
    global SESSION
    try:
        if SESSION is not None:
            SESSION.close()
    except Exception:
        pass
    SESSION = None


def _warmup():
    try:
        SESSION.get('https://auctions.yahoo.co.jp/', timeout=10, allow_redirects=True)
        time.sleep(1)
    except Exception:
        pass


def _throttled_get(url, timeout=15):
    """Thread-safe, rate-limited GET. Only one request may be in flight."""
    global LAST_REQUEST
    with REQUEST_LOCK:
        elapsed = time.time() - LAST_REQUEST
        if elapsed < MIN_DELAY:
            time.sleep(MIN_DELAY - elapsed)
        resp = get_session().get(url, timeout=timeout)
        LAST_REQUEST = time.time()
        return resp


def _looks_like_results(body):
    return body and _RESULTS_MARKER in body


def _looks_like_no_results(body):
    if not body:
        return False
    return any(m in body for m in _NO_RESULTS_MARKERS)


def fetch(url, timeout=15):
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = _throttled_get(url, timeout=timeout)
            body = resp.text
            resp.encoding = 'utf-8'
            if resp.status_code == 200 or (
                resp.status_code != 200
                and (_looks_like_results(body) or _looks_like_no_results(body))
            ):
                resp.encoding = 'utf-8'
                return resp.text
            last_err = 'HTTP {} {} for url: {}'.format(
                resp.status_code, resp.reason, resp.url
            )
        except requests.RequestException as e:
            last_err = str(e)
        if attempt < MAX_RETRIES:
            # Yahoo rate-limits burst traffic (returns 404 for blocked bots),
            # and occasionally drops connections. Back off, drop cookies,
            # and try again with a fresh session.
            time.sleep(2 * attempt)
            _reset_session()
    raise requests.HTTPError(last_err)