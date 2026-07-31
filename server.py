import os
import sys
from typing import List, Optional
from mercari import (
    MercariOrder, MercariSearchStatus, MercariSort,
    search
)
from pydantic import Field
from fastmcp import FastMCP

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'temp_download'))
from ebay_search import search_ebay, EbayItem

mercari_mcp = FastMCP(name="MercariSearchComplete")

@mercari_mcp.tool(name="search_mercari_jp",
                description="""Search Mercari for items, excluding keywords and filtering by price and specific model name.
                Args:
                    keyword (str): The main keyword to search for (e.g., 'iPhone15 Pro 256GB'). Optimize this to ensure the product name is correct, sometimes it has to be in Japanese.
                    exclude_keywords (str): Space-separated keywords to exclude. Think about exclude keywords that can make the search more precise. Generate this in japanese. For example, 'ジャンク', 'max', 'plus', '11', '12', '13', '14', '16', 'ケース', 'カバー', 'フィルム' when searching for iPhone15 Pro 256GB. Don't forget to separate them with space. Do not include '新品', '未使用', or '中古' in this list if not requested.
                    min_price (int, optional): Minimum price in JPY. Think about the minimum price that you are willing to pay for the item. For example, if you are looking for a new iPhone15 Pro 256GB, you might want to set a minimum price of 100000 JPY.
                    max_price (int, optional): Maximum price in JPY. Think about the maximum price that you are willing to pay for the item. For example, if you are looking for a new iPhone15 Pro 256GB, you might want to set a maximum price of 200000 JPY.
                    limit (int): Maximum number of items to return.""")
def search_mercari_items_filtered(
    keyword: str = Field(..., description="The main keyword to search for (e.g., 'iPhone15 Pro 256GB')."),
    exclude_keywords: str = Field("", description="Space-separated keywords to exclude (e.g., 'ジャンク max')."),
    min_price: Optional[int] = Field(None, description="Minimum price in JPY.", ge=0),
    max_price: Optional[int] = Field(None, description="Maximum price in JPY.", ge=0),
    limit: int = Field(20, description="Maximum number of items to return.", ge=1)
) -> str:
    try:
        search_results = search(
            keyword,
            sort=MercariSort.SORT_SCORE,
            order=MercariOrder.ORDER_DESC,
            status=MercariSearchStatus.ON_SALE,
            exclude_keywords=exclude_keywords,
            max_items=limit * 3
        )

        items_found: List[str] = []
        required_terms = [term.lower() for term in keyword.split()]
        unwanted_terms_from_input = [term.lower() for term in exclude_keywords.split()]
        all_unwanted_terms = list(set(unwanted_terms_from_input))
        all_unwanted_terms = [term for term in all_unwanted_terms if term not in required_terms]

        for item in search_results:
            try:
                product_name = getattr(item, 'productName', None)
                if product_name is None:
                    continue

                price = getattr(item, 'price', None)
                if price is None:
                    continue
                try:
                    price = float(price)
                except (ValueError, TypeError):
                    continue

                lower_product_name = product_name.lower()
                name_contains_desired_keywords = all(
                    term in lower_product_name for term in required_terms
                )
                name_contains_unwanted_terms = any(
                    term in lower_product_name for term in all_unwanted_terms
                )

                if name_contains_desired_keywords and not name_contains_unwanted_terms:
                    min_check_passed = (min_price is None) or (price >= min_price)
                    max_check_passed = (max_price is None) or (price <= max_price)

                    if min_check_passed and max_check_passed:
                        image_url = getattr(item, 'imageURL', 'N/A')
                        item_url = getattr(item, 'productURL', 'N/A')

                        parts = [f"### {product_name}"]
                        parts.append(f"![{product_name}]({image_url})")
                        parts.append(f"**Price:** ¥{price:,.0f} | [View on Mercari]({item_url})")

                        items_found.append("\n".join(parts))

                        if len(items_found) >= limit:
                            break

            except (AttributeError) as filter_err:
                print(f"Warning: Skipping item during post-filtering due to data access error: {filter_err}")
                continue
            except Exception as unexpected_err:
                print(f"Warning: Skipping item due to unexpected error during filtering: {unexpected_err}")
                continue
        if not items_found:
            return "No items found matching your criteria."

        return "\n\n---\n\n".join(items_found)

    except Exception as e:
        print(f"Error: An error occurred during Mercari search: {e}")
        raise e

@mercari_mcp.tool(name="search_ebay",
                description="""Search eBay for items with keyword, price, condition filtering.
                Args:
                    keyword (str): The main keyword to search for (e.g., 'iPhone 15 Pro 256GB').
                    min_price (float, optional): Minimum price in USD.
                    max_price (float, optional): Maximum price in USD.
                    condition (str, optional): Item condition - 'new', 'used', 'open_box', 'refurbished', 'for_parts'.
                    sort (str, optional): Sort order - 'best_match', 'price_asc', 'price_desc'.
                    limit (int): Maximum number of items to return (max 100).
                    global_id (str, optional): eBay site to search - 'EBAY-US', 'EBAY-GB', 'EBAY-DE', 'EBAY-JAPAN', etc.
                    bin_only (bool): Only show Buy It Now items.
                    item_location (str, optional): Filter by item location country code, e.g. 'US', 'JP', 'PK', 'GB'.
                    app_id (str, optional): eBay App ID (Client ID). Falls back to EBAY_APP_ID env var.
                    cert_id (str, optional): eBay Cert ID (Client Secret). Falls back to EBAY_CERT_ID env var.""")
def search_ebay_items(
    keyword: str = Field(..., description="The main keyword to search for (e.g., 'iPhone 15 Pro 256GB')."),
    min_price: Optional[float] = Field(None, description="Minimum price in USD.", ge=0),
    max_price: Optional[float] = Field(None, description="Maximum price in USD.", ge=0),
    condition: str = Field("", description="Item condition: new, used, open_box, refurbished, for_parts."),
    sort: str = Field("best_match", description="Sort order: best_match, price_asc, price_desc, newly_listed, ending_soon."),
    limit: int = Field(20, description="Maximum number of items to return.", ge=1, le=100),
    global_id: str = Field("EBAY-US", description="eBay site: EBAY-US, EBAY-GB, EBAY-DE, EBAY-JAPAN, etc."),
    bin_only: bool = Field(False, description="Only show Buy It Now items."),
    item_location: str = Field("", description="Filter by item location country code, e.g. 'US', 'JP', 'PK', 'GB'."),
    app_id: Optional[str] = Field(None, description="eBay App ID (Client ID). Falls back to EBAY_APP_ID env var."),
    cert_id: Optional[str] = Field(None, description="eBay Cert ID (Client Secret). Falls back to EBAY_CERT_ID env var.")
) -> str:
    try:
        ebay_app_id = app_id or os.environ.get("EBAY_APP_ID", "")
        ebay_cert_id = cert_id or os.environ.get("EBAY_CERT_ID", "")
        if not ebay_app_id:
            return "Error: eBay App ID is required. Set EBAY_APP_ID environment variable or pass app_id parameter."
        if not ebay_cert_id:
            return "Error: eBay Cert ID is required. Set EBAY_CERT_ID environment variable or pass cert_id parameter."

        results = search_ebay(
            app_id=ebay_app_id,
            cert_id=ebay_cert_id,
            keyword=keyword,
            min_price=min_price,
            max_price=max_price,
            condition=condition,
            sort=sort,
            limit=limit,
            global_id=global_id,
            bin_only=bin_only,
            item_location=item_location,
        )

        if not results:
            return f"No eBay items found for '{keyword}'."

        items_found: List[str] = []
        for item in results:
            parts = [f"### {item.title}"]
            if item.thumbnail:
                parts.append(f"![{item.title}]({item.thumbnail})")
            price_str = f"${item.price:,.2f}" if item.currency == "USD" else f"{item.currency} {item.price:,.2f}"
            details = f"**Price:** {price_str}"
            if item.buy_it_now_price:
                details += f" | BIN: ${item.buy_it_now_price:,.2f}"
            if item.condition:
                details += f" | {item.condition}"
            if item.bid_count is not None:
                details += f" | {item.bid_count} bid(s)"
            total_str = ""
            if item.free_shipping:
                details += " | Free Shipping (Total: " + price_str + ")"
            elif item.shipping_type == "CALCULATED":
                details += " | Shipping: Calculated"
            elif item.shipping_cost is not None:
                total = item.price + item.shipping_cost
                total_s = f"${total:,.2f}" if item.currency == "USD" else f"{item.currency} {total:,.2f}"
                details += f" | Shipping: ${item.shipping_cost:.2f} (Total: {total_s})"
            parts.append(details)
            parts.append(f"[View on eBay]({item.url})")
            items_found.append("\n".join(parts))

        return "\n\n---\n\n".join(items_found)

    except Exception as e:
        print(f"Error: eBay search failed: {e}")
        return f"Error searching eBay: {str(e)}"


if __name__ == "__main__":
    mercari_mcp.run()
