"""Download every published product from a Shopify store's public products.json.

Usage: python fetch_shopify.py <store-domain> <out.json>
store-domain: e.g. styllano.myshopify.com (the <handle>.myshopify.com form always works).
No login needed; only published products are returned.
"""
import json, sys, time, urllib.request

domain, out = sys.argv[1], sys.argv[2]
products, page = [], 1
while True:
    url = f'https://{domain}/products.json?limit=250&page={page}'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    batch = json.load(urllib.request.urlopen(req, timeout=60))['products']
    print(f'page {page}: {len(batch)}')
    if not batch: break
    products += batch; page += 1; time.sleep(1)
json.dump(products, open(out, 'w'), ensure_ascii=False)
print('total products:', len(products))
