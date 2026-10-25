# AJL's Lumpia and Bread Roll — POS System

Mobile-friendly POS with a Google Sheets database and a Google Apps Script API. **No login or accounts anywhere.**

```
ajls-pos/
  frontend/   index.html, config.js, config.example.js, assets/ (optional product photos)
  backend/    Code.gs, appsscript.json
  docs/       TESTING.md
```

## 1. Create the database and backend
1. Create a new Google Sheet named **AJL POS Database**.
2. **Extensions → Apps Script**. Replace the editor contents with `backend/Code.gs`.
3. Project Settings → tick **Show appsscript.json**, paste `backend/appsscript.json` (time zone Asia/Manila).
4. Select function `setup` → **Run** → approve permissions. This creates the seven sheets: Products, Sales, Sale Items, Expenses, Inventory, Settings, Audit Logs, plus Lumpia (₱5) and Bread Roll (₱10).
5. **Deploy → New deployment → Web app**. Execute as: **Me**. Who has access: **Anyone**. Copy the URL ending in `/exec`.
6. After editing `Code.gs` later, use Deploy → Manage deployments → Edit → **New version**.

## 2. Connect the frontend
Edit `frontend/config.js` and paste the URL as `API_URL`. The URL is a public endpoint, not a credential, but **anyone who has it can read and write your data** (there is no login by design). Keep it unlisted and consider a private repo.

## 3. Deploy the frontend
GitHub Pages only serves static files; the backend runs separately on Google Apps Script.
1. Push this folder to GitHub. 2. Settings → Pages → deploy from branch `main`, folder `/frontend` (or move files to `/docs`). 3. Open the Pages URL on your phone and use "Add to Home screen".
Product photos: add `frontend/assets/lumpia.jpg` and `frontend/assets/bread-roll.jpg` (emoji placeholders show until then).

## How dates and stock work
- Every sale stores **Actual Sales Date** (used by all reports) separately from **Recorded At** (when typed in). Prices are saved per sale line, so price changes never alter past sales.
- New sales deduct stock. Historical sales do **not** unless you tick the box, so stock is never deducted twice or invented. Editing/deleting a sale writes reversing "Correction" movements.
- Every create/edit/delete is written to Audit Logs. Duplicate protection: each form creates a unique ID first; the backend ignores a repeated ID.
- Profit is labelled an estimate when any sold item has no unit cost.

## Troubleshooting
- **"Backend not connected"**: `API_URL` still the placeholder.
- **"Cannot reach Google Sheets"**: wrong URL, offline, or deployment access not "Anyone". Redeploy a new version after code edits.
- **Permission prompt blocked**: Advanced → continue (your own script).
- **Dates show wrong day**: confirm the Apps Script time zone is Asia/Manila.

## Screenshots
Add screenshots to `docs/` and link them here after deployment.
