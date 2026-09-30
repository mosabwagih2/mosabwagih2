# Prompt for Claude in Chrome: fill the Vendoor color column

Paste everything below the line into Claude in Chrome (the Chrome side panel), and attach the store's color workbook (`<store>_colors.xlsx`, made by `scripts/build_color_table.py`). You must already be logged in to https://aff.ven-door.com in that Chrome.

---

You are running in my Chrome browser, and I'm logged in to Vendoor (https://aff.ven-door.com). I sell Vendoor products on my Shopify store and I'm building a unified color system. I need the color name Vendoor uses for every row of the attached workbook.

**This is a read-only task.** Only open and read product pages. Never click "أضف أوردر" (add order), add-to-cart, or anything that creates, edits or deletes something on Vendoor or Shopify.

## The file

Sheet `Products`, one row per (Shopify product × Shopify color):

| col | header | meaning |
|---|---|---|
| A | المنتج | Shopify product title |
| E | اللون في شوبيفاي | the color as written on Shopify |
| H | لينك/SKU Vendoor | the Vendoor product URL `https://aff.ven-door.com/product/<id>`. Sometimes it's an Arabic SKU text instead, or several URLs separated by ` , ` |
| I | اللون في Vendoor (يتملى) | **fill this column** |

Don't change any other column, sheet, or header.

## What to write in column I

For each row, open the Vendoor URL in column H, find the variant that matches the Shopify color in column E, and write Vendoor's color name **exactly as Vendoor writes it**. Keep the Arabic spelling. Don't translate, fix typos, or rename.

- If Vendoor's variant names include the size, keep them all and separate them with ` / `, for example `اسود XL / اسود 2XL`. The merge script strips the sizes.
- If the row already has a value (some rows get prefilled from Arabic SKUs), check it and keep it if it's right.

When there's no color to write, use one of these markers **exactly**. The merge script looks for this text:

| situation | write exactly |
|---|---|
| Vendoor page is 404 / product removed | `— المنتج اتشال من Vendoor (404)` |
| product exists but this color isn't on Vendoor | `— مش موجود على Vendoor` |
| Vendoor variants have no color, only a code, model, or letter | `— Vendoor مش كاتب لون (<what Vendoor shows>)`, e.g. `— Vendoor مش كاتب لون (M)` |

Leave column I empty only when column H has no Vendoor link at all.

## How to work

1. Open 2–3 product pages first to learn where Vendoor shows the variants and colors.
2. If you can run JavaScript in the page, `fetch('/product/<id>')` from a tab on aff.ven-door.com. It's the same origin, so my login cookie is sent. Parse each page with `DOMParser` and wait about 250 ms between requests. That's much faster than clicking through hundreds of pages. If a response redirects to `/login`, stop and tell me.
3. Otherwise, work in batches of about 25 products. Tell me the last row you finished, so I can say "continue from row N" if you stop.
4. When you're done, give me the filled workbook back as `.xlsx` (same sheets and columns) and a count of: rows filled, 404 products, colors not on Vendoor, and rows with no color on Vendoor.
