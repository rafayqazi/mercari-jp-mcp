import sys, os
sys.path.insert(0, os.path.join('temp_download', 'mercari-2.2.1'))
sys.path.insert(0, 'temp_download')

from mercari.mercari import Item

# Simulate a regular Mercari item (ITEM_TYPE_MERCARI)
regular_raw = {
    'id': 'm43988731738',
    'name': 'Test Item',
    'price': 5000,
    'status': 'ITEM_STATUS_ON_SALE',
    'created': 0,
    'updated': 0,
    'thumbnails': ['https://example.com/img.jpg'],
    'itemType': 'ITEM_TYPE_MERCARI',
    'shop': None,
    'auction': None,
}

# Simulate a Mercari Shops item (ITEM_TYPE_BEYOND with shop data)
shop_raw = {
    'id': 'Jjn7SV7mXTgU647mCrZRe7',
    'name': 'Shop Product',
    'price': 3000,
    'status': 'ITEM_STATUS_ON_SALE',
    'created': 0,
    'updated': 0,
    'thumbnails': ['https://example.com/img2.jpg'],
    'itemType': 'ITEM_TYPE_BEYOND',
    'shop': {'id': 'H5uBJNeMBKTsLMHcFsyQW3'},
    'auction': None,
}

regular = Item(**regular_raw)
shop = Item(**shop_raw)

print("=== Regular Item ===")
print(f"  id:          {regular.id}")
print(f"  isShopItem:  {regular.isShopItem}  (expected: False)")
print(f"  productURL:  {regular.productURL}  (expected: /item/)")

print("\n=== Shop Item ===")
print(f"  id:          {shop.id}")
print(f"  isShopItem:  {shop.isShopItem}  (expected: True)")
print(f"  productURL:  {shop.productURL}  (expected: /shops/product/)")

# Validate
assert not regular.isShopItem, "Regular item should NOT be a shop item"
assert 'jp.mercari.com/item/' in regular.productURL, "Regular item URL should be /item/"
assert shop.isShopItem, "Shop item SHOULD be a shop item"
assert 'jp.mercari.com/shops/product/' in shop.productURL, "Shop item URL should be /shops/product/"

print("\n✅ All assertions passed! Fix is correct.")
