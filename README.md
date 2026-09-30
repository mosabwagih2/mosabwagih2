# Shopify × Vendoor Color System

> **ملخص لمصعب:** الملف ده بيشرح بالتفصيل إزاي اتعمل الـcolor system لمتجر Styllano: سحب الألوان من شوبيفاي، وتوحيد الأسماء، ومقارنتها بـVendoor. عشان تعمل نفس الشغل لمتجر تاني، افتح شات جديد مع Claude، وادّيله الريبو ده أو الملف ده، وقوله:
> «اقرا README.md واعمل نفس الشغل لمتجر `<اسم المتجر>` — لينك الأدمن: `https://admin.shopify.com/store/<handle>/products`».

This README is written for **Claude in a new chat**. It explains what was done for the Styllano store and how to repeat it exactly for another Shopify store owned by the same person (Mosab, Egypt; he writes in Egyptian Arabic, so answer in Arabic and keep technical terms in English).

---

## 1. Goal

Mosab runs Shopify stores that resell products from **Vendoor** (https://aff.ven-door.com), an Egyptian dropshipping/affiliate supplier. Over time, color names on Shopify became inconsistent, for example `Burgundy` / `Burgandy` / `Burgany`, `Gray` / `Grey` / `رمادي`, and `Havan` / `Havana`. He wants a **unified color system**: one standard name per color, used the same way on every product.

The work has three parts:

1. **Inventory.** Pull every color value from the Shopify store and list each unique value once, with how often it's used.
2. **Standardize.** Group values that are the same color (typos, Arabic vs English, synonyms), propose one standard English name for each group, and flag the uncertain ones.
3. **Match against Vendoor.** For every Shopify product × color, find the color name Vendoor uses for that product. Then flag mismatches (wrong color, not available, product removed) and build an Arabic→English dictionary of Vendoor's color names.

The deliverables are Excel workbooks, which Mosab reviews and uses to fix products in the Shopify admin.

---

## 2. Repo layout

```
README.md                                 ← this file
scripts/
  fetch_shopify.py                        ← step 1: download products.json (public, no login)
  build_color_table.py                    ← step 2: build the color workbook
  make_vendoor_inputs.py                  ← step 3a: Vendoor link list + console grabber script
  vendoor_colors_grabber.template.js      ←   template used by make_vendoor_inputs.py
  merge_vendoor.py                        ← step 4: compare the filled Vendoor column → final workbook
prompts/
  vendoor_fill_prompt.md                  ← step 3b: prompt for Claude in Chrome to fill the Vendoor column
stores/
  styllano/                               ← everything for the Styllano store (reference run)
    products.json                         ← raw Shopify dump (620 products at fetch time)
    styllano_colors.xlsx                  ← step 2 output (Vendoor column empty)
    styllano_colors_vendoor_filled.xlsx   ← same file after Claude in Chrome filled column I
    styllano_colors_vendoor.xlsx          ← step 4 output, the final deliverable
    vendoor_matching_input.csv            ← 578 Vendoor links + Shopify colors
    vendoor_colors_grabber.js             ← console script with Styllano's 578 Vendoor ids
```

For a new store, create `stores/<handle>/` and put all of its files there.

---

## 3. Environment facts (read before starting)

These cost time in the first run. Don't rediscover them.

| Fact | Consequence |
|---|---|
| A Claude Code **cloud session cannot see or control Mosab's Chrome**. The "Browser" item in the Claude app menu is greyed out in cloud sessions. | Don't promise to "open his Chrome". Any step that needs his logged-in browser is done by him or by **Claude in Chrome** (the Chrome extension side panel he has installed). |
| Shopify's **public** `https://<handle>.myshopify.com/products.json` works without login. | Step 1 needs no credentials. It returns **published products only**. Drafts need a CSV export from the admin (Products → Export). |
| **Vendoor requires login.** `/product/<id>` redirects to `/login`, and `/api/product/<id>` returns `{"message":"Unauthenticated."}`. | Vendoor data can't be fetched from the cloud container without his session. See step 3. |
| The Vendoor link for each product is stored in the **variant SKU** field on Shopify: `https://aff.ven-door.com/product/<id>`. Most products have one link. Some have one link per color. A few (about 32 products) have an Arabic SKU text instead, like `كوتشي فلات REEBOK (مخزن 300) - ابيضXكحلي 41`, which already contains Vendoor's color name. | This is how Shopify variants are joined to Vendoor products. |
| `openpyxl` may not be installed in a fresh container. | Run `pip install openpyxl`. |
| LibreOffice (`soffice`) timed out on every file, even tiny ones, so the xlsx skill's `recalc.py` could not run. | The workbook sets `fullCalcOnLoad = True`, so Excel computes the `COUNTIF` formulas when the file opens. Verify formula ranges in Python: the sum of the per-color counts must equal the number of Products rows. |
| Headless Chromium in the container failed with `ERR_CERT_AUTHORITY_INVALID` because of the TLS-intercepting egress proxy. | Launch Playwright's Chromium with `proxy: {server: process.env.HTTPS_PROXY}` and `--ignore-certificate-errors-spki-list=<sha256 SPKI of /root/.ccr/agent-proxy-ca.crt>`. This only trusts the proxy's CA. The only thing this was used for was confirming that Vendoor redirects to login. |
| **Credentials.** Mosab offered to paste his Vendoor password. | Prefer the console script or Claude in Chrome. If he still sends it, warn him it stays in the chat history and he should change it afterwards. Never write it to a file or commit it. |

---

## 4. Step 1: fetch Shopify products

```bash
python scripts/fetch_shopify.py <handle>.myshopify.com stores/<handle>/products.json
```

- `<handle>` is the part after `/store/` in the admin URL. For `https://admin.shopify.com/store/styllano/products` it's `styllano`, so the domain is `styllano.myshopify.com`.
- The script paginates `products.json?limit=250&page=N` until it gets an empty page (Styllano: 250 + 250 + 120).
- Fields used later: `id`, `title`, `vendor`, `options[] {name, position, values}`, `variants[] {option1, option2, option3, sku}`.

**Color option detection.** The option name itself wasn't consistent. On Styllano:

| option name | products |
|---|---|
| Size | 550 |
| **Color** | 368 |
| **Colors** | 240 |
| Shoe size | 12 |
| Title | 5 (no variants) |
| **color**, **ColorB**, **الالوان**, size, المقاس | 1 each |

The scripts treat `{'Color', 'Colors', 'color', 'ColorB', 'الالوان'}` as color options (the `COLOR_OPTS` / `C` set in the scripts). **For a new store, print the option-name counter first** and add any new color-ish names to that set in `build_color_table.py` and `make_vendoor_inputs.py`.

Tell Mosab about this: Shopify's Search & Discovery builds **one filter per option name**, so `Color` and `Colors` show up as two separate filters for customers. The first fix is renaming every color option to `Color`.

---

## 5. Step 2: build the color workbook

```bash
python scripts/build_color_table.py stores/<handle>/products.json stores/<handle>/<handle>_colors.xlsx <handle>
```

### 5.1 What counts as a "row"

There's one row per **(product × value of its color option)**. Styllano had 2,032 rows and 253 distinct raw values.

### 5.2 Standardization rules (`standardize()` in the script)

1. **Normalize the raw value.** Collapse whitespace, strip a trailing `.`, and compare in lowercase.
2. **Not a color?** A regex matches `Pack…`, `Bundle…`, `AA/AB/AC/AD`, `Cat`, `Tiger`, and `Stars`. These map to `— (مش لون)`, with a note to move them to a separate option like "Pack" or "Design".
3. **Known single color.** The `S` dictionary maps a lowercase raw value to `(standard, reason)`. The reasons are:
   - `ok`: already the standard name
   - `spelling`: typo or case difference, like `Wihte`, `Brwon`, `Burgandy`, `Baige`, `Selver`, `Mouve`
   - `arabic`: Arabic written on Shopify, like `اسود`, `ابيض`, `كحلي`, `هافان`, `منجاوي`
   - `synonym`: `Navy Blue`→Navy, `Petrol Blue`→Petrol, `Kiwi Green`→Kiwi, `Off-White`→Off White
   - `review`: a plausible merge that **needs human eyes**, like `Olive Green`→Olive, `Mint`→Mint Green, `Jeans`→Denim Blue, `Simon`→Salmon, `Cold`→Gold (probably a typo), `Gold 18/20`→Gold (the number is the karat), `Matte Black`, `Black suede`, and `Havana Brown`
4. **Multi-color values.** Split on `with`, `and`, `&`, `×`, ` x `, `*`, `/`, ` – `, ` - `, and Arabic `X`. Remove detail words (`Logo`, `Stripes`, `Striped`, `Sole`, `Accent(s)`, `Mark`, `Dual-Face`, `Leather`, numbers) and record them in the note. Map each part through `S`, then join the parts with **` & `** in their original order, main color first. For example, `White with Black Logo` becomes `White & Black`.
   - Why `&` and not `/`: Shopify joins option values with ` / ` in variant titles (`Black / XL`), so `/` inside a color name is confusing.
   - Detail words belong in the description or images, not in the color name.

Chosen conventions:
- `Gray`, not `Grey`: 124 vs 30 uses. The same goes for Dark Gray and Light Gray.
- Title Case English names.
- Egyptian market names are kept where there's no clean English equivalent: `Havana` (هافان), `Cashmere` (كشمير), `Genzari` (جنزاري), `Firani`, `Pepsi` (بيبسي), `Mango` (منجاوي), `Cumin` (كموني).

### 5.3 Workbook layout

The column letters matter: `merge_vendoor.py` depends on them.

| Sheet | Contents |
|---|---|
| `اقرأني` | Legend, the color key for the status column, and conventions |
| `Standard Colors` | One row per standard name. **A** standard, **B** type (لون / لون مزدوج / مش لون), **C** `=COUNTIF(Products!F:F, A)`, **D** number of different raw spellings, **E** raw spellings with counts, **F** notes. Step 4 adds **G**, the Vendoor names. |
| `All Values` | Every raw value. **A** raw, **B** `=COUNTIF(Products!E:E, A)`, **C** suggested standard, **D** status (colored), **E** note |
| `Products` | One row per product × color. **A** title, **B** vendor, **C** admin link (`https://admin.shopify.com/store/<handle>/products/<id>`, hyperlinked "فتح"), **D** option name (highlighted red if it isn't `Color`), **E** Shopify raw color, **F** suggested standard, **G** status, **H** Vendoor URL(s) or Arabic SKU, **I** Vendoor color (yellow = to fill), **J** note |
| `Option Names` | Option-name counts and the suggested rename |

**Vendoor link per row (column H).** The script takes the SKUs of the variants whose color option equals that value. So when a product has a different Vendoor link per color, each row gets the right link. URL SKUs are joined with ` , ` (at most 3). If a SKU is Arabic text, the part after ` - ` (minus the trailing size) is pre-filled into column I.

**Checks done.** The sum of `Standard Colors!C` equals the number of Products rows (2,032). The sum of `All Values!B` also equals 2,032.

### 5.4 Styllano results (baseline)

- 620 products, 2,032 rows, and 253 raw color values, grouped into 113 standard names (66 single colors, 46 two-tone colors, and one "not a color" group).
- Biggest mess: Black (Black, اسود, أسود, Black×Black…), Gray (Gray 124 / Grey 30 / رمادي / رصاصي), Havana (**Havan 29** / Havana 24 / Havana Brown / هافان), Burgundy (20 / Burgandy 8 / Burgany 2), and Cashmere (14 / Kashmir 8 / Cashmir 1).
- About 50 two-tone colors were written in 7 separator styles.

---

## 6. Step 3: get Vendoor's color names

Vendoor needs Mosab's logged-in browser. There are three ways, in order of preference:

### 6A. Claude in Chrome fills the workbook (used for Styllano ✅)

1. Mosab opens Claude in Chrome (side panel) while logged in to aff.ven-door.com.
2. He pastes `prompts/vendoor_fill_prompt.md` and attaches `<handle>_colors.xlsx`.
3. Claude in Chrome fills **Products!I** and returns the file. For Styllano, the returned file is `stores/styllano/styllano_colors_vendoor_filled.xlsx`.

**Exact format expected in column I** (`merge_vendoor.py` parses this):
- Vendoor's own text, e.g. `زيتي`, `ابيض تطعيم اسود`, or with sizes, `اسود XL / اسود 2XL`. Parts are separated by ` / `.
- Or one of these markers, which start with `—`:
  - `— المنتج اتشال من Vendoor (404)`: the product page is gone
  - `— مش موجود على Vendoor`: the color isn't on Vendoor
  - `— Vendoor مش كاتب لون (<what it shows>)`: the variants have no color (a code, model, or letter)
- Empty only when the row has no Vendoor link.

### 6B. Console script (fast, about 3 minutes, no extension needed)

```bash
python scripts/make_vendoor_inputs.py stores/<handle>/products.json stores/<handle>/
```

This writes `vendoor_colors_grabber.js`, which has the store's Vendoor ids baked in. Mosab opens aff.ven-door.com while logged in, presses F12, opens Console, pastes the script (typing `allow pasting` first if Chrome blocks it), and presses Enter. The script `fetch`es each `/product/<id>` on the same origin with a 250 ms delay. It saves the headings, candidate option/label/table texts, and the first 20k characters of page text, then downloads `vendoor_products.json`.

**Not yet implemented:** a parser from `vendoor_products.json` into column I. Nobody on the Claude side has seen a logged-in Vendoor product page, so the selectors are unknown. If this route is used, inspect a few entries first, write the parser, and output column I in the format of 6A.

### 6C. Credentials in chat (last resort)

Only if Mosab insists. Log in with Playwright Chromium (see the proxy note in §3), loop over the ids, and never store the password. Tell him to change it afterwards.

---

## 7. Step 4: merge and compare

```bash
python scripts/merge_vendoor.py stores/<handle>/<handle>_colors_vendoor_filled.xlsx stores/<handle>/<handle>_colors_vendoor.xlsx
```

### 7.1 Normalizing Vendoor's Arabic (`ar()`)

- Letter unification: `أ إ آ → ا`, `ى → ي`, `ة → ه`, `چ → ج`, and a stray `ٍ` is removed. There's also a typo fix: `أببض → ابيض`.
- Removed: parentheses, `size-`, `N سنوات` / `٦سنوات`, `one size`, `مقاس…`, and trailing sizes (`XL`, `2XL`, `2X`, `S/M/L`, `4`, `6`, `8`…). Size removal runs twice for things like `اسود 2X L`.

### 7.2 Arabic → English dictionary (`V` in the script)

The main entries:

| Vendoor (normalized) | Standard | Notes |
|---|---|---|
| اسود، سوداء | Black | `اسود فرنيه` (finish) → Black, review |
| ابيض، بيضاء | White | |
| اوف وايت | Off White | |
| رمادي | Gray | |
| رصاصي | Gray (review) | Vendoor uses both رمادي and رصاصي. On Shopify, 21 of 23 are Gray/Grey and 2 are Dark Gray |
| رمادي غامق، رصاصي غامق | Dark Gray | |
| رمادي فاتح، رصاصي فاتح | Light Gray | |
| حديدي | Gray (review) | possibly Dark Gray |
| رمادي شانيه | Heather Gray (review) | شانيه = chiné/melange. Shopify had "Shania" |
| كحلي | Navy | |
| بيج | Beige | |
| بني | Brown | |
| ازرق | Blue | |
| هافان | Havana | |
| زيتي، زتوني | Olive | زيتي غامق → Dark Olive, زيتي فاتح → Light Olive |
| احمر، حمراء | Red | |
| جولد، ذهبي | Gold | |
| سيلفر، فضي، سلفر | Silver | |
| بينك | Pink | زهري → Pink (review) |
| روز | Rose | |
| نبيتي، برجندي | Burgundy | this is the burgundy/"عنابي" case Mosab mentioned |
| لبني، بيبي بلو | Baby Blue (review) | Shopify used Baby Blue 15× and Light Blue 7× |
| لبني فاتح | Light Blue (review) | |
| لبني غامق | Dark Baby Blue | |
| كشمير | Cashmere | |
| اصفر | Yellow | فسفوري → Neon Yellow, لموني → Lemon Yellow |
| جملي | Camel | |
| بترولي | Petrol | |
| مينت جرين، منت جرين، منت | Mint Green | |
| اخضر، خضراء | Green | |
| موف | Mauve | |
| تركواز | Turquoise | فيروزي → Turquoise (review; Shopify had Aqua, Teal, Turquoise) |
| كافيه | Coffee | |
| كاكاو | Cocoa | |
| فوشيا | Fuchsia | |
| بيبسي | Pepsi | |
| منجاوي | Mango | |
| كموني | Cumin (review) | Shopify had Camel 3× and Khaki 2× |
| جنزاري، اخضر جنزاري | Genzari (review) | Shopify had Genzari, Green, Dark Green, Teal, Petrol |
| مسترده | Mustard | |
| اورانج | Orange | |
| بستشيو | Pistachio | |
| بنفسجي | Purple | |
| صحراوي | Desert Beige (review) | رملي → Sand |
| جينز، لون جينز | Denim Blue | جينز اسود/كحلي/لبني → Black/Dark/Light Denim |
| فيراني | Firani | |
| سيمون | Salmon | |
| طوبي | Brick (review) | |
| كيوي | Kiwi | |
| خيش | Burlap | |
| بطيخي | Watermelon | |

**Not colors:** `باندل…` (bundle), `AA–AD`, `Cat`, `مشجر` (patterned). `ستارز` and `جلد` are dropped as details.

**Arabic multi-color:** split on `تطعيم`, `مطعم`, `خطوط`, `علامه`, `وعلامه`, `*`, `×`, `X` between letters, ` x `, ` في `, and `شريط`. Map each part, then join with ` & ` in Vendoor's order. For example, `ابيض تطعيم اسود` becomes `White & Black`.

**When a new store shows `❓ مش متعرف`, add the term to `V`.** Don't guess silently.

### 7.3 Status logic (Products columns K–M are added)

**K** holds the Vendoor color mapped to the standard name. If there are several distinct ones, they're joined with ` , `. **L** holds the status, and **M** holds the details.

| Status | Rule | Priority in To Fix |
|---|---|---|
| ⛔ المنتج اتشال من Vendoor | the I marker contains `404` | 1 (urgent: it's being sold but can't be fulfilled) |
| ⛔ اللون مش موجود على Vendoor | the marker is `— مش موجود` | 2 |
| ❌ مختلف | the Vendoor standard ≠ the Shopify standard (F) | 3 |
| ❓ مش متعرف | the term isn't in `V` | 4 |
| ◐ Vendoor كاتب جزء من اللون | Shopify is a two-tone color and Vendoor names only one of its colors (e.g. Shopify `White with Blue accents`, Vendoor `ازرق`) | 5 |
| ↔ نفس الألوان، ترتيب مختلف | the same set of colors in a different order | 6 |
| ✅ متطابق | equal | — |
| ℹ Vendoor مش كاتب لون / ℹ مش لون / — مفيش لينك | informational | — |

### 7.4 Output sheets added

- **`To Fix`** (second sheet): every problem row, sorted by the priority above, with admin and Vendoor links.
- **`Vendoor Dictionary`**: each normalized Vendoor term, its standard name, row count, Vendoor's exact spellings, how it was written on Shopify, and a consistency flag (⚠ = the same Vendoor color was written with several Shopify names).
- **`Standard Colors!G`**: the Vendoor Arabic names for each standard color.
- **`اقرأني`**: a summary of the counts.

### 7.5 Styllano results (baseline)

2,032 rows: **1,821 ✅**, 1 ↔, 26 ◐, **54 ❌**, **49 rows / 15 products ⛔ removed from Vendoor**, **23 rows / 19 products ⛔ color not on Vendoor**, 23 ℹ no color on Vendoor, 31 ℹ not a color, and 4 with no link. There were 153 rows in To Fix.

Notable real errors: `Baige` where Vendoor says بني (Brown), `Dark Blue` where Vendoor says زهري (Pink), `Havan` where Vendoor says جملي or بني, `Grey` where Vendoor says سلفر, `Off White` where Vendoor says ابيض, `Purple` where Vendoor says موف, and `White` where Vendoor says مشجر (patterned).

---

## 8. Open decisions (ask Mosab; don't decide for him)

| Question | Suggested |
|---|---|
| لبني = Baby Blue or Light Blue? | Baby Blue; keep Light Blue for لبني فاتح |
| رمادي vs رصاصي: one name? | Gray for both, unless رصاصي is visibly darker on his products |
| جنزاري name | Genzari |
| كموني (Shopify said Camel/Khaki) | Cumin, as its own color |
| فيروزي / تركواز | Turquoise |
| صحراوي vs رملي | Desert Beige vs Sand, or merge them |
| طوبي | Brick |
| رمادي شانيه | Heather Gray |
| Two-tone naming when Vendoor names only the accent (◐ rows) | pick one rule store-wide |

Rules for proposing merges: only merge when it's the **same shade**. If two names are different shades in front of the customer, keep both, because a wrong merge causes returns. Nobody checked the product photos. Say so, and ask him to eyeball uncertain ones.

---

## 9. Not done yet (next steps)

1. After Mosab answers §8: generate a **Shopify bulk-edit CSV** (Products → Import, "Overwrite existing products") that renames color option values to the standard names and renames the option to `Color`. Match rows by `Handle` + variant options, and test on 2–3 products first. Claude can't run the import. Mosab does it in the admin.
2. Draft or delete the 15 products removed from Vendoor, and remove the 19 unavailable colors.
3. A parser for the console-script route (§6B) if it's ever used.

---

## 10. Checklist for a new store

1. Get the admin URL from Mosab and take `<handle>` from `/store/<handle>/`.
2. `pip install openpyxl`, then `mkdir -p stores/<handle>`.
3. `python scripts/fetch_shopify.py <handle>.myshopify.com stores/<handle>/products.json`
4. Print the option-name counter and update `COLOR_OPTS` if needed (§4).
5. Check the SKU field: are Vendoor links there? If the store uses a different supplier or field, adapt `make_vendoor_inputs.py` and the column H logic.
6. `python scripts/build_color_table.py stores/<handle>/products.json stores/<handle>/<handle>_colors.xlsx <handle>`
7. Look at the `All Values` sheet for raw values that fell through to the generic multi-color path or look wrong, and extend `S` in `build_color_table.py`. Keep the Styllano names as the standard, so both stores share one color system.
8. Send the workbook plus `prompts/vendoor_fill_prompt.md` to Mosab for Claude in Chrome (§6A).
9. When the filled workbook comes back: `python scripts/merge_vendoor.py <filled.xlsx> stores/<handle>/<handle>_colors_vendoor.xlsx`. Fix every `❓` by extending `V`, then re-run.
10. Report back in Egyptian Arabic, with the numbers first. Report in this order: removed products, wrong colors, naming decisions (§8), and the rest.
11. Commit the per-store folder.
