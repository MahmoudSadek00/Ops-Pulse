"""
Ops Pulse -- reads the staging spreadsheet's 'Orders' + 'Not Shipped'
tabs (the SAME staging sheet sync_orders.py's daily job and the orders_status_native
app both write into), builds one unified per-order dataset, and computes Fulfillment /
Delivery KPIs for a chosen date range -- either one period on its own, or two periods
side by side for an "are we improving or not" comparison, plus a rule-based Summary
that calls out what got better, what got worse, and which metrics currently sit far
enough from target to count as a weak point.

Sep 2026, per Mahmoud: reads Orders + Not Shipped TOGETHER on purpose. Orders alone
undercounts Total/Cancelled/Pending Orders -- an order only gets a row in a raw
tracking sheet (and so in the staging Orders tab) once it's physically WITH the
shipping company; anything Cancelled or still Pending before that point only ever shows
up in Not Shipped (see orders_status_native/README.md and consolidation_tool's Orders
Status Check page for the full story this app is built on top of).

Pure logic, no Streamlit and no gspread CALLS made from here beyond load_orders_data()
accepting an already-authorized client -- everything else takes plain DataFrames, so it
can be unit-tested with synthetic data (same philosophy as clean.py / engine.py
elsewhere in this project family).
"""
import io
import re
import datetime as dt

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# Read-only tool -- only needs read scopes, unlike sync_orders.py / orders_status_native
# which also write to these sheets.
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
]

# Same real staging sheet baked into orders_status_native/logic.py (Aug 2026, confirmed)
# -- this tool reads the SAME staging spreadsheet, just never writes to it. app.py's
# staging_spreadsheet_id secret, if set, overrides this.
STAGING_SPREADSHEET_ID_DEFAULT = '1dZMqtqvnxe6GspH0C10AvXECB74NP-ZjDG_BihMOkmg'


def get_client(service_account_info):
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    return gspread.authorize(creds)


ORDERS_TAB = 'Orders'
NOT_SHIPPED_TAB = 'Not Shipped'

# Both tabs share this same 13-column shape at minimum (orders_status_native's own
# ORDERS_HEADER). The staging 'Orders' tab additionally carries 'Delivery Date' and
# 'Customer Phone' at the end (added Sep 2026, sync_orders.py) which 'Not Shipped' does
# not have (a Shopify "all orders" export has no shipping-company data yet) -- those 2
# just come out blank/NaT for Not Shipped rows in the unified frame, not an error.
CORE_HEADER = [
    'Source', 'Reference Number', 'Shipping Date', 'Order Date', 'Country', 'City',
    'Order Value', 'Status', 'New Customer Orders', 'Returning Customer Orders',
    'Needs Review', 'Review Reason', 'Last Synced At (UTC)',
]
EXTRA_ORDERS_COLUMNS = ['Delivery Date', 'Customer Phone']
FULL_HEADER = CORE_HEADER + EXTRA_ORDERS_COLUMNS

STATUSES = ['Delivered', 'Returned', 'Pending', 'Cancelled']

GOOGLE_SHEETS_EPOCH = dt.date(1899, 12, 30)
_ISO_DATE_RE = re.compile(r'^\s*(\d{4})-(\d{2})-(\d{2})\s*$')

# Default "how many days late counts as not-on-time" -- a placeholder until Mahmoud
# plugs in the CEO scorecard's real per-market transit windows via the UI (see app.py's
# 'On-time delivery target' inputs). Kept small and clearly a guess, not a real target.
DEFAULT_ON_TIME_TARGET_DAYS = 5

# Weak-point thresholds (Sep 2026, per Mahmoud: "sensible defaults I define, but keep
# them adjustable from the UI too" -- see app.py's threshold sliders). All in
# "how much WORSE than the baseline period counts as a weak point" terms. Rates are in
# percentage points (e.g. 2.0 = 2pp), time metrics are in days.
DEFAULT_WEAK_POINT_THRESHOLDS = {
    'delivered_rate_pp': 2.0,           # dropped by >= 2pp
    'cancelled_rate_pp': 2.0,           # rose by >= 2pp
    'returned_rate_pp': 2.0,            # rose by >= 2pp
    'pending_rate_pp': 3.0,             # rose by >= 3pp
    'fulfillment_lead_time_days': 1.0,  # got slower by >= 1 day
    'delivery_time_days': 1.0,          # got slower by >= 1 day
    'on_time_rate_pp': 5.0,             # dropped by >= 5pp
}

METRIC_LABELS = {
    'delivered_rate': 'Delivered rate',
    'returned_rate': 'Returned rate',
    'cancelled_rate': 'Cancelled rate',
    'pending_rate': 'Pending rate',
    'fulfillment_lead_time_days': 'Avg. fulfillment lead time (days)',
    'delivery_time_days': 'Avg. delivery time (days)',
    'on_time_rate': 'On-time delivery rate',
    'new_rate': 'New-customer rate',
    'returning_rate': 'Returning-customer rate',
}

# direction=+1 means "higher is better" (e.g. delivered_rate, on_time_rate);
# direction=-1 means "lower is better" (e.g. cancelled_rate, lead time).
METRIC_DIRECTION = {
    'delivered_rate': 1,
    'returned_rate': -1,
    'cancelled_rate': -1,
    'pending_rate': -1,
    'fulfillment_lead_time_days': -1,
    'delivery_time_days': -1,
    'on_time_rate': 1,
    'new_rate': 1,
    'returning_rate': 1,
}


# ---------------------------------------------------------------------------
# Reading the 2 staging tabs into one unified DataFrame
# ---------------------------------------------------------------------------

def _cell_to_date(value):
    """Both staging tabs write dates as plain ISO text ('YYYY-MM-DD') -- see
    sync_orders.py / orders_status_native's own fmt_date(). Handles that directly, plus
    a genuine Sheets date-serial number as a safety net for a cell someone later
    hand-edited into a native Sheets date type (which would come back as a serial
    number under UNFORMATTED_VALUE, not the original ISO text). Returns pd.Timestamp
    (or pd.NaT) rather than datetime.date, so date-arithmetic/.dt accessors work
    directly on a column built from this without an extra conversion step."""
    if value in (None, ''):
        return pd.NaT
    if isinstance(value, bool):
        return pd.NaT
    if isinstance(value, (int, float)):
        if 0 < value < 100000:
            return pd.Timestamp(GOOGLE_SHEETS_EPOCH) + pd.Timedelta(days=value)
        return pd.NaT
    s = str(value).strip()
    if not s:
        return pd.NaT
    m = _ISO_DATE_RE.match(s)
    if m:
        try:
            return pd.Timestamp(dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            return pd.NaT
    try:
        return pd.Timestamp(pd.to_datetime(s))
    except Exception:
        return pd.NaT


def _cell_to_number(value):
    if value in (None, ''):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def sheet_values_to_df(values, origin_tab):
    """values: raw 2D list from gspread's get_values(value_render_option=
    'UNFORMATTED_VALUE') for either tab, header row included (or []/[[]] for an empty/
    missing tab). Returns a normalized DataFrame with a fixed column set regardless of
    which tab it came from -- Not Shipped rows get NaT/None for the 2 Orders-only
    columns rather than erroring."""
    cols = FULL_HEADER + ['_origin_tab']
    if not values or len(values) < 2:
        return pd.DataFrame(columns=cols)

    header = [str(h).strip() for h in values[0]]
    body = values[1:]
    # Pad/truncate every row to the header's width -- a raw sheet read can have short
    # trailing rows (Sheets omits fully-blank trailing cells).
    width = len(header)
    body = [list(r) + [''] * (width - len(r)) if len(r) < width else r[:width] for r in body]
    raw = pd.DataFrame(body, columns=header)

    # Reindex to the full expected column set -- tolerant of a tab missing a column
    # entirely (Not Shipped has no Delivery Date/Customer Phone at all) rather than
    # raising, and tolerant of duplicate/blank header cells by just taking the first
    # match for each expected name.
    df = pd.DataFrame(index=raw.index)
    for col in FULL_HEADER:
        df[col] = raw[col] if col in raw.columns else None

    out = pd.DataFrame(index=df.index)
    out['_ref_key'] = df['Reference Number'].astype(str).str.strip().str.upper().str.lstrip('#')
    out['_source'] = df['Source'].astype(str).str.strip()
    out['_order_date'] = df['Order Date'].map(_cell_to_date)
    out['_shipping_date'] = df['Shipping Date'].map(_cell_to_date)
    out['_delivery_date'] = df['Delivery Date'].map(_cell_to_date)
    out['_country'] = df['Country'].astype(str).str.strip().str.upper()
    out['_city'] = df['City'].astype(str).str.strip()
    out['_order_value'] = df['Order Value'].map(_cell_to_number)
    out['_status'] = df['Status'].astype(str).str.strip()
    out['_new_customer'] = df['New Customer Orders'].map(_cell_to_number).fillna(0)
    out['_returning_customer'] = df['Returning Customer Orders'].map(_cell_to_number).fillna(0)
    out['_customer_phone'] = df['Customer Phone'].astype(str).str.strip()
    out.loc[out['_customer_phone'].isin(['None', 'nan']), '_customer_phone'] = ''
    out['_origin_tab'] = origin_tab
    # Drop rows with no reference number at all (fully-blank trailing rows some sheets
    # carry) -- nothing meaningful to report on without a join key.
    out = out[out['_ref_key'] != '']
    return out.reset_index(drop=True)


def load_orders_data(gc, staging_spreadsheet_id):
    """Reads both tabs live and returns one combined DataFrame, de-duplicated by
    Reference Number: an order present in BOTH tabs (the day it ships, before Not
    Shipped has been cleaned up -- see orders_status_native's "Clean up Not Shipped"
    feature) keeps only its Orders-tab row, since that one carries the real
    shipping-company data (Delivery Date, Customer Phone, the real Status) and Not
    Shipped's copy of it is stale by definition."""
    sh = gc.open_by_key(staging_spreadsheet_id)

    def _read_tab(tab_name):
        try:
            ws = sh.worksheet(tab_name)
        except Exception:
            return sheet_values_to_df([], tab_name)
        values = ws.get_values(value_render_option='UNFORMATTED_VALUE')
        return sheet_values_to_df(values, tab_name)

    orders_df = _read_tab(ORDERS_TAB)
    not_shipped_df = _read_tab(NOT_SHIPPED_TAB)
    return combine_tabs(orders_df, not_shipped_df)


def combine_tabs(orders_df, not_shipped_df):
    """Split out from load_orders_data() so it can be unit-tested without a live
    Google Sheets connection -- see the module docstring."""
    combined = pd.concat([orders_df, not_shipped_df], ignore_index=True)
    if combined.empty:
        return combined
    combined['_priority'] = (combined['_origin_tab'] == ORDERS_TAB).astype(int)
    combined = combined.sort_values('_priority', ascending=False, kind='stable')
    combined = combined.drop_duplicates(subset='_ref_key', keep='first')
    combined = combined.drop(columns=['_priority']).reset_index(drop=True)
    return combined


# ---------------------------------------------------------------------------
# Period metrics
# ---------------------------------------------------------------------------

def filter_period(df, start_date, end_date, countries=None):
    """start_date/end_date: date-like (datetime.date or pd.Timestamp), inclusive.
    Filters on Order Date -- the one date field present and populated on every row
    regardless of which tab it came from (Shipping/Delivery Date are blank for most
    Pending/Cancelled rows, which would silently drop them from a shipping-date-based
    filter and understate exactly the statuses this tool exists to surface)."""
    if df.empty:
        return df
    start_ts = pd.Timestamp(start_date)
    end_ts = pd.Timestamp(end_date)
    mask = df['_order_date'].notna() & (df['_order_date'] >= start_ts) & (df['_order_date'] <= end_ts)
    if countries:
        wanted = {c.upper() for c in countries}
        mask &= df['_country'].isin(wanted)
    return df[mask].copy()


def _avg_day_gap(sub, from_col, to_col, extra_mask=None):
    """Average (to_col - from_col) in days, only over rows where both are present and
    to_col is not before from_col (a negative gap means bad data -- a typo'd date, a
    hand-edited cell -- not a real same-day-or-earlier delivery; excluded rather than
    letting it drag the average down artificially)."""
    m = sub[from_col].notna() & sub[to_col].notna()
    if extra_mask is not None:
        m &= extra_mask
    pairs = sub[m]
    if pairs.empty:
        return None, 0
    gap_days = (pairs[to_col] - pairs[from_col]).dt.days
    gap_days = gap_days[gap_days >= 0]
    if gap_days.empty:
        return None, 0
    return float(gap_days.mean()), int(len(gap_days))


def _resolve_on_time_target(country, on_time_target_days):
    if on_time_target_days is None:
        return None
    if isinstance(on_time_target_days, dict):
        return on_time_target_days.get(country, on_time_target_days.get('_default'))
    return on_time_target_days  # single number applied to every market


def _metrics_for_slice(sub, on_time_target_days=None):
    """sub: an already-filtered (period + country) DataFrame slice. Returns the flat
    metrics dict shared by both the overall and the per-country breakdown."""
    total = len(sub)
    status_counts = {s: int((sub['_status'] == s).sum()) for s in STATUSES}
    status_rates = {s: (status_counts[s] / total if total else None) for s in STATUSES}

    fulfillment_days, fulfillment_n = _avg_day_gap(sub, '_order_date', '_shipping_date')

    delivered = sub[sub['_status'] == 'Delivered']
    delivery_days, delivery_n = _avg_day_gap(delivered, '_order_date', '_delivery_date')

    on_time_rate, on_time_n = None, 0
    if on_time_target_days is not None:
        have_both = delivered[delivered['_order_date'].notna() & delivered['_delivery_date'].notna()]
        if not have_both.empty:
            gap = (have_both['_delivery_date'] - have_both['_order_date']).dt.days
            targets = have_both['_country'].map(lambda c: _resolve_on_time_target(c, on_time_target_days))
            evaluable = targets.notna()
            if evaluable.any():
                on_time = (gap[evaluable] <= targets[evaluable])
                on_time_rate = float(on_time.mean())
                on_time_n = int(evaluable.sum())

    new_n = int(sub['_new_customer'].sum())
    returning_n = int(sub['_returning_customer'].sum())
    new_rate = new_n / total if total else None
    returning_rate = returning_n / total if total else None

    total_value = float(sub['_order_value'].sum(skipna=True)) if total else 0.0
    avg_value = (total_value / total) if total else None

    return {
        'total_orders': total,
        'status_counts': status_counts,
        'status_rates': status_rates,
        'delivered_rate': status_rates['Delivered'],
        'returned_rate': status_rates['Returned'],
        'pending_rate': status_rates['Pending'],
        'cancelled_rate': status_rates['Cancelled'],
        'fulfillment_lead_time_days': fulfillment_days,
        'fulfillment_lead_time_n': fulfillment_n,
        'delivery_time_days': delivery_days,
        'delivery_time_n': delivery_n,
        'on_time_rate': on_time_rate,
        'on_time_n': on_time_n,
        'new_rate': new_rate,
        'returning_rate': returning_rate,
        'total_order_value': total_value,
        'avg_order_value': avg_value,
    }


def compute_period_metrics(df, start_date, end_date, countries=None, on_time_target_days=None):
    """Top-level entry point for ONE period. Returns the overall metrics dict (see
    _metrics_for_slice) plus a 'per_country' dict of the same shape keyed by country
    code, and the raw filtered row count for a sanity check in the UI."""
    sub = filter_period(df, start_date, end_date, countries)
    overall = _metrics_for_slice(sub, on_time_target_days)
    per_country = {}
    if not sub.empty:
        for country, grp in sub.groupby('_country'):
            per_country[country] = _metrics_for_slice(grp, on_time_target_days)
    overall['per_country'] = per_country
    overall['start_date'] = pd.Timestamp(start_date).date()
    overall['end_date'] = pd.Timestamp(end_date).date()
    return overall


# ---------------------------------------------------------------------------
# Two-period comparison + rule-based Summary
# ---------------------------------------------------------------------------

def compare_periods(df, period_a, period_b, countries=None, on_time_target_days=None):
    """period_a / period_b: (start_date, end_date) tuples. 'a' is the baseline/earlier
    period, 'b' is what it's being compared against -- deltas are always b - a, so a
    positive delta on a "higher is better" metric (e.g. delivered_rate) means period b
    improved on period a."""
    metrics_a = compute_period_metrics(df, *period_a, countries=countries, on_time_target_days=on_time_target_days)
    metrics_b = compute_period_metrics(df, *period_b, countries=countries, on_time_target_days=on_time_target_days)

    def _delta(key, a_dict, b_dict, as_pp=False):
        va, vb = a_dict.get(key), b_dict.get(key)
        if va is None or vb is None:
            return None
        d = vb - va
        return d * 100 if as_pp else d

    rate_keys = ['delivered_rate', 'returned_rate', 'pending_rate', 'cancelled_rate', 'on_time_rate', 'new_rate', 'returning_rate']
    day_keys = ['fulfillment_lead_time_days', 'delivery_time_days']

    def _deltas_for(a_dict, b_dict):
        d = {k: _delta(k, a_dict, b_dict, as_pp=True) for k in rate_keys}
        d.update({k: _delta(k, a_dict, b_dict, as_pp=False) for k in day_keys})
        d['total_orders'] = (b_dict['total_orders'] - a_dict['total_orders']) if (a_dict['total_orders'] is not None and b_dict['total_orders'] is not None) else None
        d['total_order_value'] = b_dict['total_order_value'] - a_dict['total_order_value']
        return d

    overall_deltas = _deltas_for(metrics_a, metrics_b)
    per_country_deltas = {}
    all_countries = set(metrics_a['per_country']) | set(metrics_b['per_country'])
    for c in all_countries:
        a_c = metrics_a['per_country'].get(c)
        b_c = metrics_b['per_country'].get(c)
        if a_c is None or b_c is None:
            continue  # country only present in one period -- nothing to compare yet
        per_country_deltas[c] = _deltas_for(a_c, b_c)

    return {
        'period_a': metrics_a,
        'period_b': metrics_b,
        'deltas': overall_deltas,
        'per_country_deltas': per_country_deltas,
    }


def _fmt_pp(value):
    return f"{abs(value):.1f}pp"


def _fmt_days(value):
    return f"{abs(value):.1f}d"


def generate_summary(comparison, thresholds=None):
    """Rule-based, not AI-guessed: walks the comparison's deltas against `thresholds`
    (see DEFAULT_WEAK_POINT_THRESHOLDS) and buckets each evaluable metric into 'good'
    (moved enough in the right direction), 'weak_points' (moved enough in the WRONG
    direction -- these are what change per Mahmoud's "weak points I need to fix" ask),
    or left out of both if the move was inside the threshold (noise, not a real trend).
    Runs once for the overall numbers and once per country, so a market-specific
    problem hiding inside an OK overall number still surfaces."""
    th = dict(DEFAULT_WEAK_POINT_THRESHOLDS)
    if thresholds:
        th.update(thresholds)

    def _classify(scope, deltas):
        good, bad = [], []
        rate_metric_to_threshold = {
            'delivered_rate': 'delivered_rate_pp',
            'cancelled_rate': 'cancelled_rate_pp',
            'returned_rate': 'returned_rate_pp',
            'pending_rate': 'pending_rate_pp',
            'on_time_rate': 'on_time_rate_pp',
        }
        for metric, threshold_key in rate_metric_to_threshold.items():
            d = deltas.get(metric)
            if d is None:
                continue
            threshold = th.get(threshold_key)
            if threshold is None:
                continue
            direction = METRIC_DIRECTION[metric]
            signed = d * direction  # positive signed == moved in the GOOD direction
            word = 'rose' if d > 0 else 'dropped'
            entry = {'scope': scope, 'metric': metric, 'delta_pp': d,
                     'message': f"{METRIC_LABELS[metric]} {word} {_fmt_pp(d)} in {scope}."}
            if signed >= threshold:
                good.append(entry)
            elif signed <= -threshold:
                bad.append(entry)

        for metric, threshold_key in [('fulfillment_lead_time_days', 'fulfillment_lead_time_days'),
                                       ('delivery_time_days', 'delivery_time_days')]:
            d = deltas.get(metric)
            if d is None:
                continue
            threshold = th.get(threshold_key)
            if threshold is None:
                continue
            direction = METRIC_DIRECTION[metric]
            signed = d * direction
            word = 'got slower by' if d > 0 else 'got faster by'
            entry = {'scope': scope, 'metric': metric, 'delta_days': d,
                     'message': f"{METRIC_LABELS[metric]} {word} {_fmt_days(d)} in {scope}."}
            if signed >= threshold:
                good.append(entry)
            elif signed <= -threshold:
                bad.append(entry)
        return good, bad

    good_overall, bad_overall = _classify('overall', comparison['deltas'])
    good_countries, bad_countries = [], []
    for country, deltas in comparison['per_country_deltas'].items():
        g, b = _classify(country, deltas)
        good_countries.extend(g)
        bad_countries.extend(b)

    # Weak points: sorted worst-first so the most urgent issue leads the Summary tab.
    def _severity(item):
        return abs(item.get('delta_pp', item.get('delta_days', 0)))

    weak_points = sorted(bad_overall + bad_countries, key=_severity, reverse=True)
    good_points = sorted(good_overall + good_countries, key=_severity, reverse=True)

    return {
        'good': good_points,
        'weak_points': weak_points,
        'thresholds_used': th,
    }


# ---------------------------------------------------------------------------
# Excel export (Sep 2026, per Mahmoud: "عاوز اقدر اسحب الكلام ده اكسيل بكل
# التفاصيل و الشارتس" -- export whatever's currently on screen, full detail
# tables + charts, as one .xlsx). Charts here are NATIVE Excel chart objects
# (openpyxl.chart), built off the same tables written into the sheet -- not
# picture/image exports of the on-screen Plotly charts. That keeps them
# editable/native once opened in Excel and avoids adding an image-rendering
# dependency (kaleido) that isn't otherwise needed anywhere in this app or
# guaranteed to work on Streamlit Cloud.
# ---------------------------------------------------------------------------

_XLSX_HEADER_FILL = PatternFill(start_color='1F4E3D', end_color='1F4E3D', fill_type='solid')
_XLSX_HEADER_FONT = Font(name='Arial', bold=True, color='FFFFFF')
_XLSX_BODY_FONT = Font(name='Arial')
_XLSX_TITLE_FONT = Font(name='Arial', bold=True, size=13)
_XLSX_SUBTITLE_FONT = Font(name='Arial', italic=True, color='555555')

# (key, label, kind) -- same set/order as app.py's render_metric_cards(), kind picks
# the cell's Excel number_format so rates render as %, money as thousands-separated,
# etc. (matching how the Streamlit metric cards display them).
_METRIC_ROWS = [
    ('total_orders', 'Total Orders', 'count'),
    ('delivered_rate', 'Delivered rate', 'pct'),
    ('returned_rate', 'Returned rate', 'pct'),
    ('cancelled_rate', 'Cancelled rate', 'pct'),
    ('pending_rate', 'Pending rate', 'pct'),
    ('fulfillment_lead_time_days', 'Avg. fulfillment lead time (days)', 'days'),
    ('delivery_time_days', 'Avg. delivery time (days)', 'days'),
    ('on_time_rate', 'On-time delivery rate', 'pct'),
    ('new_rate', 'New-customer rate', 'pct'),
    ('returning_rate', 'Returning-customer rate', 'pct'),
    ('total_order_value', 'Total order value', 'money'),
    ('avg_order_value', 'Avg. order value', 'money'),
]

_PER_COUNTRY_METRICS = [
    ('total_orders', 'Total Orders', 'count'),
    ('delivered_rate', 'Delivered rate', 'pct'),
    ('returned_rate', 'Returned rate', 'pct'),
    ('cancelled_rate', 'Cancelled rate', 'pct'),
    ('pending_rate', 'Pending rate', 'pct'),
    ('fulfillment_lead_time_days', 'Fulfillment lead time (d)', 'days'),
    ('delivery_time_days', 'Delivery time (d)', 'days'),
    ('on_time_rate', 'On-time rate', 'pct'),
    ('new_rate', 'New-customer rate', 'pct'),
    ('returning_rate', 'Returning-customer rate', 'pct'),
    ('total_order_value', 'Total order value', 'money'),
]

_COMPARISON_RATE_KEYS = [
    ('delivered_rate', 'Delivered rate'),
    ('cancelled_rate', 'Cancelled rate'),
    ('returned_rate', 'Returned rate'),
    ('pending_rate', 'Pending rate'),
]

_NUMBER_FORMATS = {'pct': '0.0%', 'days': '0.0"d"', 'money': '#,##0', 'count': '#,##0'}


def _xlsx_bytes(wb):
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


def _write_title_block(ws, lines, start_row=1):
    r = start_row
    for i, line in enumerate(lines):
        cell = ws.cell(row=r, column=1, value=line)
        cell.font = _XLSX_TITLE_FONT if i == 0 else _XLSX_SUBTITLE_FONT
        r += 1
    return r + 1  # first free row, with a blank-row gap


def _write_header_row(ws, row, left_col, headers):
    for c, name in enumerate(headers, start=left_col):
        cell = ws.cell(row=row, column=c, value=name)
        cell.font = _XLSX_HEADER_FONT
        cell.fill = _XLSX_HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions[get_column_letter(c)].width = max(16, len(str(name)) + 2)


def _write_metric_value_table(ws, top_row, left_col, metric_defs, metrics):
    """Metric | Value table -- one row per (key, label, kind) in metric_defs, value
    pulled from metrics[key]. Returns (first_data_row, last_data_row)."""
    _write_header_row(ws, top_row, left_col, ['Metric', 'Value'])
    r = top_row + 1
    for key, label, kind in metric_defs:
        ws.cell(row=r, column=left_col, value=label).font = _XLSX_BODY_FONT
        val_cell = ws.cell(row=r, column=left_col + 1, value=metrics.get(key))
        val_cell.font = _XLSX_BODY_FONT
        val_cell.number_format = _NUMBER_FORMATS[kind]
        r += 1
    ws.column_dimensions[get_column_letter(left_col)].width = 36
    return top_row + 1, r - 1


def _write_comparison_metric_table(ws, top_row, left_col, metrics_a, metrics_b, deltas):
    """Metric | Period A | Period B | Delta | Unit -- same rows as app.py's Comparison
    tab 'Overall deltas' table."""
    _write_header_row(ws, top_row, left_col, ['Metric', 'Period A', 'Period B', 'Delta', 'Unit'])
    r = top_row + 1
    for key, label in METRIC_LABELS.items():
        d = deltas.get(key)
        if d is None:
            continue
        kind = 'pct' if key.endswith('_rate') else 'days'
        unit = 'pp' if kind == 'pct' else 'days'
        ws.cell(row=r, column=left_col, value=label).font = _XLSX_BODY_FONT
        for offset, val in ((1, metrics_a.get(key)), (2, metrics_b.get(key))):
            cell = ws.cell(row=r, column=left_col + offset, value=val)
            cell.font = _XLSX_BODY_FONT
            cell.number_format = _NUMBER_FORMATS[kind]
        delta_cell = ws.cell(row=r, column=left_col + 3, value=d / 100 if kind == 'pct' else d)
        delta_cell.font = _XLSX_BODY_FONT
        delta_cell.number_format = '+0.0%;-0.0%' if kind == 'pct' else '+0.0"d";-0.0"d"'
        ws.cell(row=r, column=left_col + 4, value=unit).font = _XLSX_BODY_FONT
        r += 1
    ws.column_dimensions[get_column_letter(left_col)].width = 34
    return top_row + 1, r - 1


def _write_status_table_and_chart(ws, top_row, left_col, metrics, title, anchor_col):
    headers = ['Status', 'Count']
    _write_header_row(ws, top_row, left_col, headers)
    r = top_row + 1
    for s in STATUSES:
        ws.cell(row=r, column=left_col, value=s).font = _XLSX_BODY_FONT
        cell = ws.cell(row=r, column=left_col + 1, value=metrics['status_counts'][s])
        cell.font = _XLSX_BODY_FONT
        cell.number_format = _NUMBER_FORMATS['count']
        r += 1
    last_row = r - 1

    chart = BarChart()
    chart.type = 'bar'
    chart.title = title
    chart.legend = None
    chart.height, chart.width = 7, 12
    data = Reference(ws, min_col=left_col + 1, min_row=top_row, max_row=last_row)
    cats = Reference(ws, min_col=left_col, min_row=top_row + 1, max_row=last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, f"{get_column_letter(anchor_col)}{top_row}")
    return last_row


def _write_per_country_sheet(ws, per_country, title):
    ws.cell(row=1, column=1, value=title).font = _XLSX_TITLE_FONT
    top_row = 3
    headers = ['Country'] + [label for _, label, _ in _PER_COUNTRY_METRICS]
    _write_header_row(ws, top_row, 1, headers)
    r = top_row + 1
    for country in sorted(per_country):
        m = per_country[country]
        ws.cell(row=r, column=1, value=country).font = _XLSX_BODY_FONT
        for c, (key, _, kind) in enumerate(_PER_COUNTRY_METRICS, start=2):
            cell = ws.cell(row=r, column=c, value=m.get(key))
            cell.font = _XLSX_BODY_FONT
            cell.number_format = _NUMBER_FORMATS[kind]
        r += 1
    last_row = r - 1
    if last_row < top_row + 1:
        return  # no countries in this slice -- nothing to chart

    col_of = {key: c for c, (key, _, _) in enumerate(_PER_COUNTRY_METRICS, start=2)}
    chart_row = last_row + 3
    anchor_col = 1
    for key, label in [('delivered_rate', 'Delivered rate'), ('cancelled_rate', 'Cancelled rate'),
                        ('returned_rate', 'Returned rate')]:
        chart = BarChart()
        chart.type = 'col'
        chart.title = f"{label} by country"
        chart.legend = None
        chart.y_axis.numFmt = '0%'
        chart.height, chart.width = 7, 11
        data = Reference(ws, min_col=col_of[key], min_row=top_row, max_row=last_row)
        cats = Reference(ws, min_col=1, min_row=top_row + 1, max_row=last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"{get_column_letter(anchor_col)}{chart_row}")
        anchor_col += 6


def export_single_period_xlsx(metrics, meta):
    """metrics: compute_period_metrics()'s return value. meta: dict with 'countries'
    (list), 'on_time_target_days', 'generated_at' (str). Returns .xlsx bytes with the
    same numbers/breakdown as the Overview tab in "Single period" mode."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Overview'
    r = _write_title_block(ws, [
        'Ops Pulse -- Single period report',
        f"Order Date {metrics['start_date']} → {metrics['end_date']}",
        f"Countries: {', '.join(meta['countries'])}",
        f"On-time delivery target: {meta['on_time_target_days']} day(s)",
        f"Generated: {meta['generated_at']}",
    ])
    _, last = _write_metric_value_table(ws, r, 1, _METRIC_ROWS, metrics)
    _write_status_table_and_chart(ws, last + 3, 1, metrics, 'Orders by status', anchor_col=4)

    ws2 = wb.create_sheet('Per Country')
    _write_per_country_sheet(ws2, metrics['per_country'], f"{metrics['start_date']} → {metrics['end_date']} -- by country")

    return _xlsx_bytes(wb)


def export_comparison_xlsx(comparison, summary, meta):
    """comparison: compare_periods()'s return value. summary: generate_summary()'s
    return value. meta: dict with 'countries', 'on_time_target_days', 'generated_at'.
    Returns .xlsx bytes covering the Overview + Comparison + Summary tabs in "Compare
    two periods" mode."""
    metrics_a, metrics_b, deltas = comparison['period_a'], comparison['period_b'], comparison['deltas']

    wb = Workbook()
    ws = wb.active
    ws.title = 'Overview'
    r = _write_title_block(ws, [
        'Ops Pulse -- Comparison report',
        f"Period A (baseline): {metrics_a['start_date']} → {metrics_a['end_date']}",
        f"Period B: {metrics_b['start_date']} → {metrics_b['end_date']}",
        f"Countries: {', '.join(meta['countries'])}",
        f"On-time delivery target: {meta['on_time_target_days']} day(s)",
        f"Generated: {meta['generated_at']}",
    ])
    _, last = _write_comparison_metric_table(ws, r, 1, metrics_a, metrics_b, deltas)
    chart_row = last + 3
    last_a = _write_status_table_and_chart(ws, chart_row, 1, metrics_a, 'Period A -- by status', anchor_col=4)
    _write_status_table_and_chart(ws, chart_row, 8, metrics_b, 'Period B -- by status', anchor_col=11)

    ws_a = wb.create_sheet('Per Country - Period A')
    _write_per_country_sheet(ws_a, metrics_a['per_country'], f"Period A: {metrics_a['start_date']} → {metrics_a['end_date']}")
    ws_b = wb.create_sheet('Per Country - Period B')
    _write_per_country_sheet(ws_b, metrics_b['per_country'], f"Period B: {metrics_b['start_date']} → {metrics_b['end_date']}")

    ws_cmp = wb.create_sheet('Comparison')
    ws_cmp.cell(row=1, column=1, value='Period A (baseline) vs. Period B, by market').font = _XLSX_TITLE_FONT
    row_cursor = 3
    for rate_key, label in _COMPARISON_RATE_KEYS:
        countries_here = sorted(set(metrics_a['per_country']) | set(metrics_b['per_country']))
        rows = []
        for c in countries_here:
            va = metrics_a['per_country'].get(c, {}).get(rate_key)
            vb = metrics_b['per_country'].get(c, {}).get(rate_key)
            if va is None and vb is None:
                continue
            rows.append((c, va, vb))
        if not rows:
            continue
        ws_cmp.cell(row=row_cursor, column=1, value=f"{label} by country").font = _XLSX_SUBTITLE_FONT
        table_row = row_cursor + 1
        _write_header_row(ws_cmp, table_row, 1, ['Country', 'Period A', 'Period B'])
        r2 = table_row + 1
        for c, va, vb in rows:
            ws_cmp.cell(row=r2, column=1, value=c).font = _XLSX_BODY_FONT
            for col_offset, val in ((2, va), (3, vb)):
                cell = ws_cmp.cell(row=r2, column=col_offset, value=val)
                cell.font = _XLSX_BODY_FONT
                cell.number_format = _NUMBER_FORMATS['pct']
            r2 += 1
        last_row2 = r2 - 1

        chart = BarChart()
        chart.type = 'col'
        chart.title = f"{label} by country — A vs B"
        chart.height, chart.width = 7, 12
        chart.y_axis.numFmt = '0%'
        data = Reference(ws_cmp, min_col=2, max_col=3, min_row=table_row, max_row=last_row2)
        cats = Reference(ws_cmp, min_col=1, min_row=table_row + 1, max_row=last_row2)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws_cmp.add_chart(chart, f"F{row_cursor}")

        row_cursor = max(last_row2, row_cursor + 15) + 3

    ws_sum = wb.create_sheet('Summary')
    ws_sum.cell(row=1, column=1, value='✅ What\'s working').font = _XLSX_TITLE_FONT
    r = 3
    _write_header_row(ws_sum, r, 1, ['Scope', 'Metric', 'Message'])
    r += 1
    if summary['good']:
        for item in summary['good']:
            ws_sum.cell(row=r, column=1, value=item['scope']).font = _XLSX_BODY_FONT
            ws_sum.cell(row=r, column=2, value=METRIC_LABELS.get(item['metric'], item['metric'])).font = _XLSX_BODY_FONT
            ws_sum.cell(row=r, column=3, value=item['message']).font = _XLSX_BODY_FONT
            r += 1
    else:
        ws_sum.cell(row=r, column=1, value='Nothing moved enough (given the thresholds used) to call it a clear improvement.').font = _XLSX_BODY_FONT
        r += 1

    r += 2
    ws_sum.cell(row=r, column=1, value='⚠️ Weak points').font = _XLSX_TITLE_FONT
    r += 2
    _write_header_row(ws_sum, r, 1, ['Scope', 'Metric', 'Message'])
    r += 1
    if summary['weak_points']:
        for item in summary['weak_points']:
            ws_sum.cell(row=r, column=1, value=item['scope']).font = _XLSX_BODY_FONT
            ws_sum.cell(row=r, column=2, value=METRIC_LABELS.get(item['metric'], item['metric'])).font = _XLSX_BODY_FONT
            ws_sum.cell(row=r, column=3, value=item['message']).font = _XLSX_BODY_FONT
            r += 1
    else:
        ws_sum.cell(row=r, column=1, value='No weak points crossed the thresholds in this comparison.').font = _XLSX_BODY_FONT
        r += 1

    r += 2
    ws_sum.cell(row=r, column=1, value='Thresholds used').font = _XLSX_TITLE_FONT
    r += 2
    _write_header_row(ws_sum, r, 1, ['Threshold', 'Value'])
    r += 1
    for key, val in summary['thresholds_used'].items():
        ws_sum.cell(row=r, column=1, value=key).font = _XLSX_BODY_FONT
        ws_sum.cell(row=r, column=2, value=val).font = _XLSX_BODY_FONT
        r += 1
    ws_sum.column_dimensions['A'].width = 14
    ws_sum.column_dimensions['B'].width = 30
    ws_sum.column_dimensions['C'].width = 80

    return _xlsx_bytes(wb)
