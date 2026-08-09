from dataclasses import dataclass, field
from typing import Optional

@dataclass
class EbayItem:
    id: str
    title: str
    price: float
    currency: str = 'USD'
    condition: str = ''
    condition_id: Optional[int] = None
    thumbnail: str = ''
    url: str = ''
    listing_type: str = ''
    bid_count: Optional[int] = None
    buy_it_now_price: Optional[float] = None
    seller_name: str = ''
    seller_rating: str = ''
    seller_feedback_percent: str = ''
    location: str = ''
    shipping_cost: Optional[float] = None
    shipping_type: str = ''
    free_shipping: bool = False
    end_time: str = ''
    is_auction: bool = False
    top_rated_seller: bool = False
    category_name: str = ''
