# Task: match Styllano's Shopify colors against Vendoor

You are running in my Chrome browser. I am already logged in to Vendoor (https://aff.ven-door.com). I sell Vendoor products on my Shopify store, Styllano. I'm building a unified color system, and I need to know what each product's colors are called on Vendoor compared with what I wrote on Shopify.

**Read-only task.** Open and read product pages only. Do not click "Add order", "Add to cart" or anything that creates, edits or deletes something on Vendoor or Shopify.

## Input

The attached file `vendoor_matching_input.csv` has 578 rows, one per Vendoor product:

| column | meaning |
|---|---|
| `vendoor_id` | Vendoor product id |
| `vendoor_url` | `https://aff.ven-door.com/product/<id>` |
| `shopify_product` | product title on my Shopify store |
| `shopify_colors` | the color values I used on Shopify for this Vendoor product, separated by ` \| ` |

Example rows:

```
6054,https://aff.ven-door.com/product/6054,Women's high-waist leggings,Black & Tiger
6052,https://aff.ven-door.com/product/6052,Rosaline Wide-Leg Trousers,Navy | Grey | Beige | Brown | Burgundy
```

## What to do for each row

1. Open `vendoor_url`.
2. Find every color that Vendoor lists for that product (color selector, variant/attribute table, or product description). Copy them **exactly** as Vendoor writes them, whether Arabic or English.
3. Compare them with `shopify_colors`, color by color.

### Faster method (use it if you can run JavaScript in the page)

Don't click through 578 pages. From a tab on `aff.ven-door.com`:

1. Open 2–3 product pages normally and find where the colors live in the HTML (for example `select option`, a variants table, or elements with a `color` class).
2. Write one JavaScript loop that `fetch`es `/product/<id>` for every id (same origin, so my login cookie is sent), parses each page with `DOMParser`, and pulls the colors with that selector.
3. Wait about 250 ms between requests.
4. If a response redirects to `/login`, stop and tell me.

### If you have to browse page by page

Work in batches of 25. After each batch, output that batch's CSV rows, so progress isn't lost if you stop. Continue until all 578 rows are done, and tell me the last `vendoor_id` you finished.

## Output

Return one CSV with exactly these columns, one row per Shopify color (plus one row for each Vendoor color that has no Shopify match):

```
vendoor_id,shopify_product,shopify_color,vendoor_color,status,note
```

For `status`, use one of:

- `MATCH`: the same color with the same name (ignore case and extra spaces).
- `SAME_COLOR_DIFFERENT_NAME`: clearly the same color, written differently. Examples: `Grey`/`Gray`, `Burgandy`/`Burgundy`, `اسود`/`Black`, `Havan`/`هافان`.
- `NOT_ON_VENDOOR`: the color is on my Shopify but Vendoor doesn't have it.
- `ONLY_ON_VENDOOR`: Vendoor has the color but my Shopify doesn't.
- `UNSURE`: you can't tell whether they're the same color. Say why in `note`.
- `PAGE_ERROR`: the page didn't load, the product was removed, or it shows no colors.

Rules:

- Don't guess. If you're not sure, use `UNSURE`.
- Keep Vendoor's text exactly as written.
- Also put these in `note`: out-of-stock colors (if Vendoor shows stock), and any product whose Vendoor title looks like a different product from `shopify_product`.

## Reference: my proposed standard color names

Black, White, Off White, Gray, Dark Gray, Light Gray, Beige, Navy, Brown, Olive, Havana, Blue, Dark Blue, Light Blue, Baby Blue, Dark Baby Blue, Red, Burgundy, Gold, Silver, Pink, Light Pink, Rose, Dusty Rose, Mint Green, Cashmere, Camel, Yellow, Neon Yellow, Petrol, Green, Dark Green, Mauve, Coffee, Cocoa, Fuchsia, Turquoise, Teal, Aqua, Denim Blue, Light Denim, Dark Denim, Black Denim, Mango, Purple, Pepsi, Mustard, Orange, Pistachio, Khaki, Kiwi, Sand, Desert Beige, Salmon, Brick, Terracotta, Taupe, Peach, Camouflage Green, Dark Olive, Light Olive, Genzari, Firani, Cumin, Burlap, Shania.

Two-tone colors use the format `Black & White`, with the main color first.

When you're done, give me the CSV plus a short summary: how many rows are `MATCH`, `SAME_COLOR_DIFFERENT_NAME`, `NOT_ON_VENDOOR`, `ONLY_ON_VENDOOR`, `UNSURE` and `PAGE_ERROR`.
