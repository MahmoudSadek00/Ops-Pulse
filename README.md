# Ops Pulse (standalone tool)

A separate, standalone Streamlit app -- its own link, its own deployment (Sep 2026, per
Mahmoud: same pattern as the `orders_status_native` app -- kept apart on purpose).

**What it's for:** a Fulfillment/Delivery KPI report read straight off the staging
spreadsheet -- Total Orders, Delivered/Returned/Pending/Cancelled rates, fulfillment
lead time, delivery time, on-time delivery rate, new-vs-returning split, per country
-- for any date range you pick. Read-only: this app never writes anything back.

**The "are we improving or not" part:** instead of picking one date range, switch to
**Compare two periods** in the sidebar and pick a baseline (Period A) and a comparison
period (Period B). Every number above then shows side by side with a delta, there's a
dedicated **Comparison** tab with A-vs-B charts per market, and a **Summary** tab that
auto-generates a plain-English readout of what got better, what got worse, and which
metrics currently count as a **weak point** -- a rule-based readout (not guessed), using
adjustable thresholds in the sidebar (defaults are provided, but every threshold has its
own input if you want tighter or looser ones).

**Reads BOTH the staging "Orders" tab AND the "Not Shipped" tab, together (Sep 2026, per
Mahmoud).** Orders alone undercounts Total/Cancelled/Pending Orders -- an order only
gets a row in a raw tracking sheet (and so in the staging Orders tab) once it's
physically WITH the shipping company; anything Cancelled, or still Pending before that
point, only ever shows up in Not Shipped (the holding tab the separate
`orders_status_native` app writes into). An order that happens to sit in both tabs at
once (the day it ships, before Not Shipped gets cleaned up) keeps only its Orders-tab
row, since that one carries the real shipping-company data.

**On-time delivery rate** needs a target ("delivered within how many days counts as
on-time") to mean anything -- there's a placeholder (5 days, applied to every market
the same) in the sidebar's Advanced settings until the CEO scorecard's real per-market
transit windows are typed in there instead.

**Download as Excel (Sep 2026, per Mahmoud).** A "⬇️ Download this report as Excel"
button sits above the tabs -- it exports exactly what's currently on screen (whichever
mode/filters/date range are picked in the sidebar) as one `.xlsx` file:
- **Single period mode:** an "Overview" sheet (every KPI card's number + the by-status
  breakdown) and a "Per Country" sheet (the same KPIs broken out per market).
- **Compare two periods mode:** "Overview" (Period A vs. B vs. Delta table), "Per
  Country - Period A" / "- Period B", a "Comparison" sheet (A-vs-B by market, one table
  per rate), and a "Summary" sheet (the same What's working / Weak points readout as
  the Summary tab, plus the thresholds used).

The charts in the workbook are **native Excel chart objects** built from the tables
right next to them (not picture exports of the on-screen Plotly charts) -- open,
editable, and re-colorable in Excel like any chart you'd build there yourself.

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Setup: deploying this as its own app

### Part 1 -- Google Sheets access

This app needs the SAME "robot" service account already set up for the daily sync
(`orders-sync` project) and the `orders_status_native` app -- no new Google account or
extra sharing needed. This app only ever READS the staging sheet (never writes), so the
Viewer access that account already has is enough on its own.

If you don't have that service account's `.json` key file handy anymore, generate a
fresh one for the same account: Google Cloud Console -> IAM & Admin -> Service Accounts
-> (the existing `orders-sync-bot` account) -> Keys tab -> Add Key -> Create new key ->
JSON.

### Part 2 -- Put the code on GitHub

1. Go to https://github.com and sign in (same account used for the other tools).
2. **+ -> New repository**. Name it e.g. `ops-pulse`. Keep it **Private**
   (recommended). Click **Create repository**.
3. Upload every file from this folder: `app.py`, `logic.py`, `requirements.txt`,
   `README.md` (drag-and-drop, or "uploading an existing file"). Commit.

### Part 3 -- Deploy on Streamlit Cloud

1. Go to https://share.streamlit.io and sign in (same account used for the other
   Streamlit apps).
2. Click **New app** -> pick the `ops-pulse` repository, branch `main`, main
   file path `app.py`. Click **Deploy**.
3. Once it's up, you'll have a NEW, SEPARATE link for this app (its own
   `*.streamlit.app` URL) -- different from every other tool's link.

### Part 4 -- Add the secret

The staging sheet ID is already baked into `logic.py` (the same real staging sheet the
other tools use) -- the only secret actually needed is the Google credential.

1. On this app's page (share.streamlit.io), click the **⋮** menu (top-right) ->
   **Settings** -> **Secrets**.
2. Paste in the following:

   ```toml
   gcp_service_account_json = '''
   PASTE_THE_FULL_CONTENT_OF_THE_SERVICE_ACCOUNT_JSON_KEY_FILE_HERE
   '''
   ```

   Open the `.json` key file (Part 1 above) with a text editor, select all, and paste
   its exact content between the `'''` lines, unchanged. **Important:** keep the `'''`
   (three single quotes) exactly as shown, not `"""` (three double quotes) --
   otherwise Streamlit "helpfully" converts the `\n` sequences inside the key text and
   corrupts it.

   (Optional: if the staging sheet ever moves, add
   `staging_spreadsheet_id = "THE_NEW_ID"` above the JSON block to override the
   built-in default -- not needed otherwise.)
3. Click **Save**. The app restarts automatically and should now work with no further
   setup needed.

## How the numbers are computed

- **Every date filter is against Order Date** -- the one date field populated on every
  row regardless of which tab it came from (Shipping Date/Delivery Date are blank for
  most Pending/Cancelled rows, so filtering on those would silently drop exactly the
  orders this tool exists to surface).
- **Fulfillment lead time** = Shipping Date − Order Date, averaged over orders that have
  both (any source).
- **Delivery time** = Delivery Date − Order Date, averaged over **Delivered** orders
  that have both. Delivery Date currently only exists for UAE & Oman and Gulf (see the
  `orders-sync` project's own README) -- Iraq orders are excluded from this average
  until that's captured there too, not silently defaulted to 0.
- **On-time delivery rate** = share of Delivered orders whose Delivery Date − Order Date
  is within the configured target (sidebar). An order in a market with no target
  configured is left out of the rate rather than guessed.
- **A negative day-gap (a bad/typo'd date) is excluded from every average**, not let to
  drag it down artificially.
- **The Summary tab's Weak Points** are a straightforward threshold check on Period B
  vs Period A's deltas -- e.g. "Cancelled rate rose 5.3pp" fires once the rise crosses
  the sidebar's "Cancelled rate rise (pp)" threshold. Every metric is checked both
  overall and per market, so a market-specific problem hiding inside an OK overall
  number still surfaces.

## Files

- **`logic.py`** -- all the data-reading/metrics/comparison/summary logic, no
  Streamlit dependency (pure Python + pandas + gspread). Self-contained on purpose --
  reads the same staging sheet as the other tools but has zero code dependency on any
  of them.
- **`app.py`** -- the Streamlit UI: filters, metric cards, charts (Plotly), the
  Comparison tab, and the auto-generated Summary tab.
