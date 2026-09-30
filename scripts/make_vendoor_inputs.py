"""From a products.json dump, write the Vendoor link list and the browser-console grabber.

Usage: python make_vendoor_inputs.py <all.json> <out-dir>
Writes <out-dir>/vendoor_matching_input.csv and <out-dir>/vendoor_colors_grabber.js.
Vendoor links are read from each variant's SKU (https://aff.ven-door.com/product/<id>).
"""
import csv, json, os, re, sys

SRC, OUT = sys.argv[1], sys.argv[2]
COLOR_OPTS = {'Color', 'Colors', 'color', 'ColorB', 'الالوان'}
prods = json.load(open(SRC))
m = {}
for p in prods:
    copt = [o for o in p['options'] if o['name'] in COLOR_OPTS]
    key = 'option%d' % copt[0]['position'] if copt else None
    for v in p['variants']:
        mm = re.search(r'ven-door\.com/product/(\d+)', v.get('sku') or '')
        if not mm: continue
        e = m.setdefault(mm.group(1), {'titles': [], 'colors': []})
        if p['title'] not in e['titles']: e['titles'].append(p['title'])
        c = v.get(key) if key else ''
        if c and c not in e['colors']: e['colors'].append(c)

with open(os.path.join(OUT, 'vendoor_matching_input.csv'), 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f); w.writerow(['vendoor_id', 'vendoor_url', 'shopify_product', 'shopify_colors'])
    for vid, e in m.items():
        w.writerow([vid, 'https://aff.ven-door.com/product/' + vid, ' || '.join(e['titles']), ' | '.join(e['colors'])])

js = open(os.path.join(os.path.dirname(__file__), 'vendoor_colors_grabber.template.js')).read()
open(os.path.join(OUT, 'vendoor_colors_grabber.js'), 'w').write(js.replace('__IDS__', json.dumps([int(i) for i in m])))
print('vendoor products:', len(m))
