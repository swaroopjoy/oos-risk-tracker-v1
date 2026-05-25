# Stock OOS Risk Tracker

Identify out-of-stock scenarios across overlapping promotions.
Stock is scoped per **Seller + Country + Base SKU** — the same SKU
in different countries is treated as completely independent stock.

## Features
- Upload `.xlsx` reservation tracker sheet
- Auto-parse combo SKUs (`A+Bx6+Cx2`)
- Conflict detection within same seller+country scope only
- OOS risk heatmap calendar (D-3 to D+7)
- Restock panel: Hard OOS (locked) + Soft risk (locked + nominated)
- Export to formatted Excel (4 sheets)
- Filter by Seller, Country, Brand, Channel, Campaign type, Stock lock

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push this repo to GitHub
2. Go to https://share.streamlit.io
3. Click **New app** → select your repo → set `app.py` as entry point
4. Click **Deploy**

## Expected sheet columns

| Column | Required |
|---|---|
| Seller code | Yes |
| Country | Yes |
| Brand | Yes |
| Channel | Yes |
| SKU | Yes |
| Campaign / Promotion Name | Yes |
| Campaign Type | Yes |
| Stock Lock / Reserved | Yes (Yes/No) |
| Promo Start Date | Yes |
| Promo End Date | Yes |
| Today's Stock for Base SKU | Yes |
| Reserved by DKSH | Optional |
| Reserved by Graas | Optional |
| Reserved by MP | Optional |
| Nominated stock (Non Reservation) | Optional |

## SKU format

- `101336151x8` → 8 units of base SKU 101336151
- `101336151+101336152x6` → 1 unit of 101336151 + 6 units of 101336152
- `Ax2+Bx3+C` → 2 units of A + 3 units of B + 1 unit of C
