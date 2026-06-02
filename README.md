# 랩밴드 / 메디켓 발주 관리 시스템

Next.js web app for calculating and generating purchase orders for medical bandage supplies.
Deployed to Vercel (serverless). No database — all state lives in browser `localStorage`.

---

## Key constraints before writing code

- **Read `AGENTS.md` first** — this Next.js version has breaking changes from common training data.
- **No file system on Vercel** — the Excel template is embedded as a base64 string in `lib/orderTemplate.js`. If the template file changes, regenerate it with `node scripts/encode-template.js` and commit the updated `lib/orderTemplate.js`.
- **Two Excel libraries** — `xlsx` (SheetJS) is used **client-side only** for reading uploaded files; `exceljs` is used **server-side only** for generating the order Excel. Do not swap them — `exceljs` preserves merged cells and styles; `xlsx` would destroy them.

---

## File structure

```
lib/
  products.js        ← Single source of truth for all product/size data (edit here to add products)
  orderTemplate.js   ← Base64-encoded Excel order form template (렙메디케어 sheet)
  usageData.js       ← Legacy; superseded by products.js, kept for reference only

pages/
  index.js           ← Entire UI (React). All components are in this one file.
  api/
    generate-order.js ← POST endpoint: receives orders[], writes them into the Excel template

styles/
  Home.module.css    ← All styles
```

---

## How to add a new product

**All product configuration lives in `lib/products.js`.** Adding a new product requires only editing that file — no changes to `pages/index.js` or the API route.

### Calculated product (stock-based, like Mediquet / Rapband)

Append to `PRODUCTS` in `lib/products.js`:

```js
{
  id: 'newproduct',          // unique slug, used as state key and in API payload
  name: '표시이름',           // displayed in UI and summary
  type: 'calculated',
  color: '#hex',             // section header accent color
  excelName: 'ExcelName',   // text written to 품명 column in the order Excel
  hasExcelUpload: true,      // show a file picker in the Upload section
  defaultOcha: { S: 0, M: 0, L: 0 },  // default warehouse-vs-system discrepancy per size
  sizes: [
    {
      id: 'S',               // size ID — must match what appears as (S) in Excel sheet names
      label: 'S 레드',        // shown in size chips
      chipColor: '#hex',
      excelSize: 'S(레드)',   // written to 규격 column in the order Excel
      avg_monthly: 50.0,
      monthly: { '2026-03': 45, '2026-04': 52, '2026-05': 53 },
      default_stock: 80,
    },
    // ... more sizes
  ],
},
```

### Manual-order product (user types qty directly, like 스타킹)

```js
{
  id: 'newmanual',
  name: '표시이름',
  type: 'manual',
  color: '#hex',
  excelName: 'ExcelName',
  hasExcelUpload: false,
  sizes: [
    { id: 'L', label: 'L', chipColor: '#hex', excelSize: 'L', default_stock: 0 },
  ],
},
```

That's it. The UI components (`OchaPanel`, `UploadSection`, `ProductTable`, `ManualProductSection`, summary) and the API route all iterate `PRODUCTS` from the registry — no further edits needed.

---

## Data flow

```
Excel upload (client)
  → xlsx reads D3 cell from each sheet whose name contains (SIZE)
  → returns { sizeId: jeonsanQty }
  → handleUpload: stock = jeonsan + ocha (discrepancy)
  → stocks state updated, jeonsan state updated, adjustments reset

Stock input (manual edit)
  → stocks state updated directly
  → adjustments for that product reset

Order calculation (pure, per size)
  safetyStock = ceil(avg_monthly)          ← 1× monthly average
  shortage    = safetyStock - currentStock
  packs       = ceil(shortage / 30)
  qty         = packs × 30

Odd-pack panel
  → user clicks 줄이기/늘리기
  → adjustments[product][size] ±= 1 pack

Excel download
  → POST /api/generate-order  { orders: [{ product, size, qty }] }
  → server loads template, sets D4=today, clears rows 10-22
  → writes groups (one blank row between products)
  → streams .xlsx back as attachment
```

---

## localStorage schema

Key: `bandOrderStocks`

```json
{
  "stocks":       { "mediquet": { "M1": 52 }, "rapband": { "M1": 84 } },
  "jeonsan":      { "mediquet": { "M1": null }, "rapband": { "M1": null } },
  "adjustments":  { "mediquet": { "M1": 0 }, "rapband": { "M1": 0 } },
  "ocha":         { "mediquet": { "M1": 22 }, "rapband": { "M1": 3 } },
  "manualOrders": { "stocking": { "L": 0, "XL": 0 } }
}
```

State is saved on every change and restored on page load. Unrecognized keys are ignored, so adding new products is backward-compatible with existing saved state.

---

## Excel template

- Sheet name: `렙메디케어`
- Date cell: `D4` (set to today's date on generation)
- Order rows: `A10:H22` (13 rows maximum before overflow)
- Columns written per row: `A` (번호), `D` (품명), `F` (규격), `H` (수량)
- All other cells, borders, merged cells, and styles are left untouched

---

## Business rules

| Rule | Value |
|---|---|
| Pack size | 30 units |
| Box size | 60 units = 2 packs |
| Free shipping threshold | 6 boxes total (calculated products combined) |
| Safety stock multiplier | 1× monthly average |
| Order unit | Ceiling to next full pack |
| Odd-pack constraint | Must resolve before download (줄이기 / 늘리기) |

---

## Deployment

Vercel. Push to `main` triggers automatic deploy.
No environment variables required. No backend database.
