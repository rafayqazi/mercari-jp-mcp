from dataclasses import dataclass, field
from typing import Optional

@dataclass
class YahooAuctionItem:
    id: str
    title: str
    price: int
    buy_now_price: Optional[int] = None
    bid_count: int = 0
    end_time: Optional[str] = None
    end_timestamp: Optional[int] = None
    time_remaining: str = ''
    seller_id: str = ''
    seller_name: Optional[str] = None
    thumbnail: str = ''
    url: str = ''
    free_shipping: bool = False
    unused: bool = False
    is_new: bool = False
    status: str = 'live'  # live or sold
