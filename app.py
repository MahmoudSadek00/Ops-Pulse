import json
import datetime as dt

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio

from logic import (
    STATUSES, METRIC_LABELS, METRIC_DIRECTION,
    DEFAULT_DELIVERY_WINDOWS, DEFAULT_KPI_BANDS, DEFAULT_NET_DELIVERY_MATURED_THRESHOLD,
    DEFAULT_WEAK_POINT_THRESHOLDS, DEFAULT_USD_RATES, STAGING_SPREADSHEET_ID_DEFAULT,
    get_client, load_orders_data, compute_period_metrics, compare_periods, generate_summary,
    export_single_period_xlsx, export_comparison_xlsx, classify_band, classify_delivery_time_band,
    per_country_kpi_rows, convert_order_values_to_usd,
)

st.set_page_config(page_title="Ops Pulse", layout="wide")

# Sep 2026, per Mahmoud: dark professional theme -- the actual color palette lives in
# .streamlit/config.toml ([theme] base="dark" + accent/background colors); this CSS
# block only dresses up a few widgets config.toml can't reach on its own (metric cards
# as bordered "tiles", section-header underlines, a bordered look for tables and the
# download button). Every selector below is a stable Streamlit data-testid, not a class
# name that shifts between versions -- if a future Streamlit upgrade ever renames one,
# that one rule just stops applying (harmless) rather than breaking the page.
st.markdown("""
<style>
    div[data-testid="stMetric"] {
        background: linear-gradient(180deg, #171A21 0%, #14161B 100%);
        border: 1px solid #262B35;
        border-radius: 12px;
        padding: 14px 18px 10px 18px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.35);
    }
    div[data-testid="stMetricLabel"] p {
        color: #9AA4B2 !important;
        font-size: 0.82rem !important;
        font-weight: 500 !important;
        letter-spacing: .02em;
        text-transform: uppercase;
    }
    div[data-testid="stMetricValue"] { color: #F4F6F8; font-weight: 700; }
    h1, h2, h3 { letter-spacing: -0.01em; }
    h2, h3 { border-bottom: 1px solid #262B35; padding-bottom: 6px; margin-top: 1.6rem; }
    div[data-testid="stDataFrame"] { border: 1px solid #262B35; border-radius: 10px; overflow: hidden; }
    section[data-testid="stSidebar"] { border-right: 1px solid #262B35; }
    div[data-testid="stDownloadButton"] button {
        background-color: #5B8DEF; color: white; border: none; font-weight: 600;
        border-radius: 8px;
    }
    div[data-testid="stDownloadButton"] button:hover { background-color: #4879D9; color: white; }
</style>
""", unsafe_allow_html=True)

# So every Plotly chart on the page (status breakdown, per-country bars, comparison
# bars) matches the app's own dark background by default instead of clashing with a
# white plot area -- set once, globally, rather than repeating a template= kwarg on
# every px.bar/go.Figure call below.
pio.templates.default = "plotly_dark"

st.title("Ops Pulse")
st.caption(
    "Reads the staging spreadsheet's \"Orders\" + \"Not Shipped\" tabs TOGETHER (Sep "
    "2026, per Mahmoud) -- Orders alone undercounts Total/Cancelled/Pending Orders, "
    "since an order only gets a row in a raw tracking sheet once it's physically WITH "
    "the shipping company; anything Cancelled or still Pending before that only ever "
    "shows up in Not Shipped. Read-only -- this tool never writes to either tab."
)

# Fixed, non-cycled colors -- one per Status/period identity, reused consistently
# across every chart on this page (dataviz principle: color follows the entity, not
# chart-by-chart defaults).
STATUS_COLORS = {
    'Delivered': '#2E7D32',   # green
    'Returned': '#F9A825',    # amber
    'Pending': '#5C6BC0',     # indigo
    'Cancelled': '#C62828',   # red
}
PERIOD_A_COLOR = '#5C6BC0'
PERIOD_B_COLOR = '#2E7D32'

ALL_COUNTRIES = ['UAE', 'OM', 'SA', 'KW', 'QA', 'IQ']
# Markets the CEO scorecard has a transit window for -- Iraq is deliberately absent
# (see DEFAULT_DELIVERY_WINDOWS in logic.py) until Delivery Date is captured there.
WINDOWED_COUNTRIES = ['UAE', 'OM', 'SA', 'QA', 'KW']

BAND_EMOJI = {'below': '🔴', 'target': '🟢', 'exceed': '🟣'}
BAND_TEXT = {'below': 'Below target', 'target': 'On target', 'exceed': 'Exceeding target'}


def _band_caption(band):
    """band: 'below'/'target'/'exceed'/None (see logic.classify_band /
    classify_delivery_time_band). Returns a short emoji + text label, or None."""
    if band is None:
        return None
    return f"{BAND_EMOJI[band]} {BAND_TEXT[band]}"


# ---------------------------------------------------------------------------
# Auth (same pattern as orders_status_native's app.py -- same staging sheet, same
# service-account secret, read-only scopes only)
# ---------------------------------------------------------------------------

def _load_creds_info():
    try:
        if 'gcp_service_account' in st.secrets:
            return dict(st.secrets['gcp_service_account'])
        if 'gcp_service_account_json' in st.secrets:
            raw = st.secrets['gcp_service_account_json']
            return json.loads(raw) if isinstance(raw, str) else dict(raw)
    except Exception:
        pass
    return None


def _load_staging_id():
    try:
        return st.secrets.get('staging_spreadsheet_id') or STAGING_SPREADSHEET_ID_DEFAULT
    except Exception:
        return STAGING_SPREADSHEET_ID_DEFAULT


creds_info = _load_creds_info()
staging_id = _load_staging_id()

if not creds_info or not staging_id:
    st.error(
        "This app needs Google Sheets access configured once first -- see the README's "
        "one-time setup section for exactly what to paste into this app's "
        "Settings -> Secrets on Streamlit Cloud."
    )
    st.stop()


@st.cache_resource(show_spinner=False)
def _client():
    return get_client(creds_info)


try:
    gc = _client()
except Exception as e:
    st.error(f"Couldn't connect to Google Sheets with the configured credential: {e}")
    st.stop()


@st.cache_data(ttl=600, show_spinner="Reading \"Orders\" + \"Not Shipped\"...")
def _load_data(_gc, staging_spreadsheet_id, cache_bump):
    # _gc is underscore-prefixed so Streamlit doesn't try to hash the gspread client
    # (unhashable); cache_bump has NO underscore on purpose -- it IS part of the cache
    # key, so incrementing it in session_state (the sidebar "Refresh" button) forces a
    # fresh read even before the 600s TTL expires.
    return load_orders_data(_gc, staging_spreadsheet_id)


if 'cache_bump' not in st.session_state:
    st.session_state['cache_bump'] = 0

with st.sidebar:
    st.header("Data")
    if st.button("🔄 Refresh from Google Sheets"):
        st.session_state['cache_bump'] += 1
        st.cache_data.clear()

try:
    df = _load_data(gc, staging_id, st.session_state['cache_bump'])
except Exception as e:
    st.error(f"Couldn't read the staging sheet: {e}")
    st.stop()

if df.empty:
    st.warning(
        "Both \"Orders\" and \"Not Shipped\" came back empty (or don't exist yet) on "
        "the staging sheet -- nothing to report on."
    )
    st.stop()

data_min_date = df['_order_date'].min().date()
data_max_date = df['_order_date'].max().date()

with st.sidebar:
    st.caption(f"{len(df):,} orders loaded -- Order Date {data_min_date} to {data_max_date}.")

    st.header("Filters")
    countries = st.multiselect("Country / market", ALL_COUNTRIES, default=ALL_COUNTRIES)

    mode = st.radio("Mode", ["Single period", "Compare two periods"], index=1)

    def _require_range(value, label):
        # A real Streamlit date-range picker returns a single date (not a 2-tuple)
        # while the user has only picked the start date and hasn't picked the end
        # date yet -- unpacking that directly would crash with a ValueError. Stop
        # gracefully instead and ask for the second date.
        if not isinstance(value, (tuple, list)) or len(value) != 2:
            st.info(f"Pick both a start and end date for {label}.")
            st.stop()
        return value

    if mode == "Single period":
        default_start = max(data_min_date, data_max_date - dt.timedelta(days=29))
        start_date, end_date = _require_range(st.date_input(
            "Date range (Order Date)", value=(default_start, data_max_date),
            min_value=data_min_date, max_value=data_max_date,
        ), "the date range")
    else:
        st.subheader("Period B (compared against A)")
        default_b_start = max(data_min_date, data_max_date - dt.timedelta(days=29))
        b_start, b_end = _require_range(st.date_input(
            "Period B date range", value=(default_b_start, data_max_date),
            min_value=data_min_date, max_value=data_max_date, key="period_b",
        ), "Period B")
        st.subheader("Period A (baseline)")
        period_len = (b_end - b_start).days
        default_a_end = max(data_min_date, b_start - dt.timedelta(days=1))
        default_a_start = max(data_min_date, default_a_end - dt.timedelta(days=period_len))
        a_start, a_end = _require_range(st.date_input(
            "Period A date range", value=(default_a_start, default_a_end),
            min_value=data_min_date, max_value=data_max_date, key="period_a",
        ), "Period A")
        if (b_end - b_start).days != (a_end - a_start).days:
            st.caption(
                "⚠️ Period A and B are different lengths -- compare the RATES (%) "
                "below, not the raw counts, since counts alone aren't scale-comparable."
            )

    with st.expander("Advanced settings"):
        st.caption(
            "Per-market delivery windows (days, Shipping Date → Delivery Date) -- "
            "defaults are the CEO Q3 2026 scorecard's own numbers. The High day doubles "
            "as On-Time Delivery Rate's target for that market (delivered within it = "
            "on-time); Low is only used for Delivery Time's own Below/Target/Exceed "
            "band. Iraq has no window yet (Delivery Date isn't captured there) -- "
            "excluded from On-Time Delivery Rate and Delivery Time bands until it is."
        )
        delivery_windows = {}
        for c in WINDOWED_COUNTRIES:
            lo_default, hi_default = DEFAULT_DELIVERY_WINDOWS[c]
            wcol1, wcol2 = st.columns(2)
            lo = wcol1.number_input(f"{c} -- Low (d)", min_value=1, max_value=60, value=lo_default, key=f"win_lo_{c}")
            hi = wcol2.number_input(f"{c} -- High (d)", min_value=lo, max_value=60, value=max(hi_default, lo), key=f"win_hi_{c}")
            delivery_windows[c] = (lo, hi)
        on_time_target_days = {c: hi for c, (lo, hi) in delivery_windows.items()}

        st.caption(
            "Net Delivery Rate's \"matured cohort\" gate (CEO scorecard: \"no rate "
            "quoted while >10% in transit\") -- the share of a period's Shipped orders "
            "still Pending (in transit) has to be at or under this before the rate is "
            "quoted at all; otherwise it's flagged as not-yet-matured instead of shown."
        )
        net_delivery_matured_threshold = st.number_input(
            "Max share still in transit to quote the rate (%)",
            min_value=1.0, max_value=100.0, value=DEFAULT_NET_DELIVERY_MATURED_THRESHOLD * 100, step=1.0,
        ) / 100.0

        st.caption("Below / Target / Exceed cutoffs (CEO Q3 2026 scorecard) -- everything under Target counts as Below.")
        bc1, bc2 = st.columns(2)
        on_time_target_pct = bc1.number_input("On-time rate -- Target (%)", 50.0, 100.0, DEFAULT_KPI_BANDS['on_time_rate']['target'] * 100, 0.5) / 100.0
        on_time_exceed_pct = bc2.number_input("On-time rate -- Exceed (%)", 50.0, 100.0, DEFAULT_KPI_BANDS['on_time_rate']['exceed'] * 100, 0.5) / 100.0
        nc1, nc2 = st.columns(2)
        net_delivery_target_pct = nc1.number_input("Net delivery rate -- Target (%)", 50.0, 100.0, DEFAULT_KPI_BANDS['net_delivery_rate']['target'] * 100, 0.5) / 100.0
        net_delivery_exceed_pct = nc2.number_input("Net delivery rate -- Exceed (%)", 50.0, 100.0, DEFAULT_KPI_BANDS['net_delivery_rate']['exceed'] * 100, 0.5) / 100.0

        st.caption("Weak-point thresholds -- how much worse than the baseline period counts as a weak point.")
        thresholds = {}
        thresholds['delivered_rate_pp'] = st.number_input("Delivered rate drop (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['delivered_rate_pp'], 0.5)
        thresholds['cancelled_rate_pp'] = st.number_input("Cancelled rate rise (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['cancelled_rate_pp'], 0.5)
        thresholds['returned_rate_pp'] = st.number_input("Returned rate rise (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['returned_rate_pp'], 0.5)
        thresholds['pending_rate_pp'] = st.number_input("Pending rate rise (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['pending_rate_pp'], 0.5)
        thresholds['fulfillment_lead_time_days'] = st.number_input("Fulfillment lead time slowdown (days)", 0.1, 10.0, DEFAULT_WEAK_POINT_THRESHOLDS['fulfillment_lead_time_days'], 0.1)
        thresholds['delivery_time_days'] = st.number_input("Delivery time slowdown (days)", 0.1, 10.0, DEFAULT_WEAK_POINT_THRESHOLDS['delivery_time_days'], 0.1)
        thresholds['on_time_rate_pp'] = st.number_input("On-time rate drop (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['on_time_rate_pp'], 0.5)
        thresholds['net_delivery_rate_pp'] = st.number_input("Net delivery rate drop (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['net_delivery_rate_pp'], 0.5)

        # Sep 2026, per Mahmoud: every money figure on the staging sheet is recorded in
        # its OWN market's native currency (Iraq in IQD, UAE & Oman in AED, the whole
        # Gulf group in BHD -- see logic.COUNTRY_CURRENCY), never converted anywhere
        # today. "USD" here converts every money figure (cards + per-country table +
        # xlsx download) to USD using the rates below; "Original" leaves everything
        # exactly as recorded, same as before this setting existed.
        st.caption("Currency for every money figure on this page and in the download.")
        currency_choice = st.radio("Currency", ["Original (as recorded)", "USD"], horizontal=True)
        currency_mode = 'USD' if currency_choice == 'USD' else 'original'
        usd_rates = dict(DEFAULT_USD_RATES)
        if currency_mode == 'USD':
            st.caption(
                "1 USD = how many of each currency -- NOT live rates, starting "
                "defaults only, edit freely. AED and BHD are both long-standing hard "
                "pegs to the Dollar; IQD drifts more, so it's worth checking that one "
                "periodically."
            )
            rc1, rc2, rc3 = st.columns(3)
            usd_rates['AED'] = rc1.number_input("1 USD = _ AED", min_value=0.0001, value=DEFAULT_USD_RATES['AED'], step=0.0001, format="%.4f")
            usd_rates['BHD'] = rc2.number_input("1 USD = _ BHD", min_value=0.0001, value=DEFAULT_USD_RATES['BHD'], step=0.0001, format="%.4f")
            usd_rates['IQD'] = rc3.number_input("1 USD = _ IQD", min_value=0.0001, value=DEFAULT_USD_RATES['IQD'], step=1.0, format="%.2f")

# Sep 2026: convert EVERY order's value to USD once here, right after the currency
# setting is known and before any period filtering/metric computation happens -- every
# metric downstream (cards, per-country table, xlsx export) just reads _order_value the
# same way it always has, so this single conversion point is what keeps the screen and
# the download in sync automatically (see logic.convert_order_values_to_usd).
unmapped_currency_countries = set()
if currency_mode == 'USD':
    df, unmapped_currency_countries = convert_order_values_to_usd(df, usd_rates)
if unmapped_currency_countries:
    st.warning(
        f"No currency mapped for: {', '.join(sorted(unmapped_currency_countries))} -- "
        "their Order Value was left UNCONVERTED (still in whatever currency it was "
        "recorded in), not guessed at."
    )

if not countries:
    st.warning("Pick at least one country/market in the sidebar.")
    st.stop()


# ---------------------------------------------------------------------------
# Small render helpers
# ---------------------------------------------------------------------------

def _pct(x):
    return f"{x * 100:.1f}%" if x is not None else "—"


def _days(x):
    return f"{x:.1f}d" if x is not None else "—"


def _money(x):
    # Sep 2026: USD mode shows a $ prefix + 2 decimals (a converted figure is smaller
    # and less round than the native-currency one, so the extra precision matters more
    # here); Original mode is unchanged from before this setting existed.
    if x is None:
        return "—"
    if currency_mode == 'USD':
        return f"${x:,.2f}"
    return f"{x:,.0f}"


def render_metric_cards(metrics, delta_metrics=None):
    """delta_metrics: same-shaped dict of deltas (period B vs A) to show as a Streamlit
    st.metric delta -- color direction picked from METRIC_DIRECTION so 'higher is
    better' and 'lower is better' metrics both show green-for-good/red-for-bad
    automatically, never a color that means the opposite of what happened."""
    def _dc(metric_key):
        return "normal" if METRIC_DIRECTION.get(metric_key, 1) == 1 else "inverse"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Orders", f"{metrics['total_orders']:,}",
               delta=(f"{delta_metrics['total_orders']:+,}" if delta_metrics and delta_metrics.get('total_orders') is not None else None),
               delta_color="off")
    c2.metric("Delivered rate", _pct(metrics['delivered_rate']),
               delta=(f"{delta_metrics['delivered_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('delivered_rate') is not None else None),
               delta_color=_dc('delivered_rate'))
    c3.metric("Returned rate", _pct(metrics['returned_rate']),
               delta=(f"{delta_metrics['returned_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('returned_rate') is not None else None),
               delta_color=_dc('returned_rate'))
    c4.metric("Cancelled rate", _pct(metrics['cancelled_rate']),
               delta=(f"{delta_metrics['cancelled_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('cancelled_rate') is not None else None),
               delta_color=_dc('cancelled_rate'))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Pending rate", _pct(metrics['pending_rate']),
               delta=(f"{delta_metrics['pending_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('pending_rate') is not None else None),
               delta_color=_dc('pending_rate'))
    c6.metric("Avg. fulfillment lead time", _days(metrics['fulfillment_lead_time_days']),
               delta=(f"{delta_metrics['fulfillment_lead_time_days']:+.1f}d" if delta_metrics and delta_metrics.get('fulfillment_lead_time_days') is not None else None),
               delta_color=_dc('fulfillment_lead_time_days'))
    c7.metric("Avg. delivery time", _days(metrics['delivery_time_days']),
               delta=(f"{delta_metrics['delivery_time_days']:+.1f}d" if delta_metrics and delta_metrics.get('delivery_time_days') is not None else None),
               delta_color=_dc('delivery_time_days'),
               help=f"Shipping Date -> Delivery Date (the shipping leg only, not time spent before it shipped). n={metrics['delivery_time_n']:,} Delivered orders with both dates.")
    # Delivery Time's band is per-market (the window varies by country -- see
    # DEFAULT_DELIVERY_WINDOWS), so it's only meaningful here when exactly one market is
    # selected; with several markets mixed together, the average can't be banded against
    # any single window without being misleading. The per-country table below always
    # shows it market-by-market regardless.
    if len(countries) == 1 and countries[0] in delivery_windows:
        dt_band = classify_delivery_time_band(metrics['delivery_time_days'], delivery_windows[countries[0]])
        cap = _band_caption(dt_band)
        if cap:
            c7.caption(cap)
    c8.metric("On-time delivery rate", _pct(metrics['on_time_rate']),
               delta=(f"{delta_metrics['on_time_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('on_time_rate') is not None else None),
               delta_color=_dc('on_time_rate'),
               help=f"Against each market's own transit window. n={metrics['on_time_n']:,}.")
    cap = _band_caption(classify_band(metrics['on_time_rate'], on_time_target_pct, on_time_exceed_pct))
    if cap:
        c8.caption(cap)

    # Sep 2026: st.columns(2) here (vs. columns(4) everywhere else on this row) is what
    # was making this row look full of blank space -- each of the 2 cards sat in a box
    # TWICE as wide as every other card on the page, with all the extra width reading as
    # empty gap since st.metric() doesn't stretch its content to fill it. columns(4)
    # with only the first 2 slots used matches every other row's card width exactly, so
    # the leftover space shrinks down to the same small gaps the rest of the grid has.
    c8b, c8c, _c8d, _c8e = st.columns(4)
    if metrics.get('net_delivery_matured'):
        c8b.metric("Net delivery rate", _pct(metrics['net_delivery_rate']),
                    delta=(f"{delta_metrics['net_delivery_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('net_delivery_rate') is not None else None),
                    delta_color=_dc('net_delivery_rate'),
                    help="(Delivered − Returned) ÷ Shipped, matured cohorts only. "
                         f"n={metrics['shipped_n']:,} Shipped orders.")
        cap = _band_caption(classify_band(metrics['net_delivery_rate'], net_delivery_target_pct, net_delivery_exceed_pct))
        if cap:
            c8b.caption(cap)
    else:
        share = metrics.get('in_transit_share')
        c8b.metric("Net delivery rate", "—")
        c8b.caption(
            f"⏳ Not matured yet -- {share*100:.0f}% of Shipped orders still in transit."
            if share is not None else "⏳ No Shipped orders in this period yet."
        )
    c8c.metric("Shipped orders", f"{metrics['shipped_n']:,}", help="Orders with a Shipping Date -- reached the shipping company, regardless of outcome yet.")
    if metrics.get('in_transit_share') is not None:
        c8c.caption(f"{metrics['in_transit_share']*100:.1f}% still in transit (Pending)")

    c9, c10, c11, c12 = st.columns(4)
    c9.metric("New-customer rate", _pct(metrics['new_rate']),
               delta=(f"{delta_metrics['new_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('new_rate') is not None else None),
               delta_color=_dc('new_rate'))
    c10.metric("Returning-customer rate", _pct(metrics['returning_rate']),
               delta=(f"{delta_metrics['returning_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('returning_rate') is not None else None),
               delta_color=_dc('returning_rate'))
    c11.metric("Total order value", _money(metrics['total_order_value']),
               delta=(f"{delta_metrics['total_order_value']:+,.0f}" if delta_metrics and delta_metrics.get('total_order_value') is not None else None),
               delta_color="off")
    c12.metric("Avg. order value", _money(metrics['avg_order_value']))

    # The money behind each status (Sep 2026, per Mahmoud: "عاوز احسب فلوس الاوردرات
    # الوصلت و الرجعت و الاتكنسل والبندج") -- same card concept as every metric above,
    # not a separate table. Delivered value is "higher is better" like delivered_rate;
    # Returned/Cancelled/Pending value are "lower is better" like their rates.
    c13, c14, c15, c16 = st.columns(4)
    for col, status, direction in (
        (c13, 'Delivered', 1), (c14, 'Returned', -1), (c15, 'Cancelled', -1), (c16, 'Pending', -1),
    ):
        delta_key = f'status_value_{status}'
        col.metric(
            f"{status} value", _money(metrics['status_value'][status]),
            delta=(f"{delta_metrics[delta_key]:+,.0f}" if delta_metrics and delta_metrics.get(delta_key) is not None else None),
            delta_color=("normal" if direction == 1 else "inverse"),
        )


def status_breakdown_chart(metrics, title):
    counts = metrics['status_counts']
    fig = go.Figure(go.Bar(
        x=[counts[s] for s in STATUSES], y=STATUSES, orientation='h',
        marker_color=[STATUS_COLORS[s] for s in STATUSES],
        text=[f"{counts[s]:,}" for s in STATUSES], textposition='outside',
        hovertemplate='%{y}: %{x:,}<extra></extra>',
    ))
    fig.update_layout(title=title, xaxis_title="Orders", yaxis_title=None,
                       height=280, margin=dict(l=10, r=10, t=40, b=10))
    return fig


def per_country_rate_chart(metrics, rate_key, title):
    rows = [{'Country': c, 'value': m[rate_key]} for c, m in metrics['per_country'].items() if m[rate_key] is not None]
    if not rows:
        return None
    d = pd.DataFrame(rows).sort_values('value', ascending=False)
    fig = px.bar(d, x='Country', y='value', title=title, text=d['value'].map(lambda v: f"{v*100:.1f}%"))
    fig.update_traces(marker_color=STATUS_COLORS.get(title.split()[0], '#5C6BC0'), textposition='outside')
    fig.update_layout(yaxis_tickformat='.0%', yaxis_title=None, xaxis_title=None,
                       height=320, margin=dict(l=10, r=10, t=40, b=10))
    return fig


def per_country_kpi_table(metrics):
    """One row per market with Net Delivery Rate / On-Time Delivery Rate / Delivery
    Time, each next to its Below/Target/Exceed band -- the on-screen equivalent of the
    CEO scorecard's 3 logistics KPI rows, market by market (Sep 2026). Built from
    logic.per_country_kpi_rows() -- the SAME function the xlsx export's "Logistics
    KPIs" sheet calls, so the table on screen and the one in the download can never
    silently drift apart; this just adds the emoji captions/formatting for display."""
    raw_rows = per_country_kpi_rows(
        metrics, delivery_windows,
        (on_time_target_pct, on_time_exceed_pct),
        (net_delivery_target_pct, net_delivery_exceed_pct),
    )
    rows = []
    for row in raw_rows:
        window = row['delivery_window']
        if row['net_delivery_matured']:
            nd_display = _pct(row['net_delivery_rate'])
        elif row.get('shipped_n'):
            nd_display = f"⏳ {(row.get('in_transit_share') or 0)*100:.0f}% in transit"
        else:
            nd_display = "—"
        rows.append({
            'Country': row['country'],
            'Net delivery rate': nd_display,
            'ND band': _band_caption(row['net_delivery_band']) or '—',
            'On-time rate': _pct(row['on_time_rate']),
            'OT band': _band_caption(row['on_time_band']) or '—',
            'Delivery time': _days(row['delivery_time_days']),
            'Window': f"{window[0]}–{window[1]}d" if window else '—',
            'DT band': _band_caption(row['delivery_time_band']) or '—',
        })
    return pd.DataFrame(rows)


def comparison_bar_chart(comparison, rate_key, title, as_pct=True):
    countries_here = sorted(set(comparison['period_a']['per_country']) | set(comparison['period_b']['per_country']))
    rows = []
    for c in countries_here:
        va = comparison['period_a']['per_country'].get(c, {}).get(rate_key)
        vb = comparison['period_b']['per_country'].get(c, {}).get(rate_key)
        if va is not None:
            rows.append({'Country': c, 'Period': 'Period A', 'value': va})
        if vb is not None:
            rows.append({'Country': c, 'Period': 'Period B', 'value': vb})
    if not rows:
        return None
    d = pd.DataFrame(rows)
    fig = px.bar(
        d, x='Country', y='value', color='Period', barmode='group', title=title,
        color_discrete_map={'Period A': PERIOD_A_COLOR, 'Period B': PERIOD_B_COLOR},
        category_orders={'Period': ['Period A', 'Period B']},
    )
    if as_pct:
        fig.update_layout(yaxis_tickformat='.0%')
    fig.update_layout(yaxis_title=None, xaxis_title=None, height=340, margin=dict(l=10, r=10, t=40, b=10))
    return fig


# ---------------------------------------------------------------------------
# Compute + render
# ---------------------------------------------------------------------------

export_meta = {
    'countries': countries,
    'on_time_target_days': on_time_target_days,
    # Sep 2026: so the downloaded .xlsx carries the same Below/Target/Exceed bands (and
    # the "Logistics KPIs" sheet) shown on screen, not just the raw numbers -- see
    # logic.py's export_single_period_xlsx / export_comparison_xlsx / per_country_kpi_rows.
    'delivery_windows': delivery_windows,
    'on_time_bands': (on_time_target_pct, on_time_exceed_pct),
    'net_delivery_bands': (net_delivery_target_pct, net_delivery_exceed_pct),
    # Sep 2026: so every money cell in the download gets the same $ formatting as the
    # on-screen cards -- see logic.py's _number_formats / convert_order_values_to_usd.
    'currency_mode': currency_mode,
    'generated_at': dt.datetime.now().strftime('%Y-%m-%d %H:%M'),
}

# Single-country selections can band Delivery Time (and the "Overall deltas" table's
# Delivery Time row, in comparison mode) against that one market's own window -- with
# several markets mixed together the average can't be banded against any one window.
_single_country = countries[0] if len(countries) == 1 else None


def _band_for_metric(key, value):
    if key == 'on_time_rate':
        return _band_caption(classify_band(value, on_time_target_pct, on_time_exceed_pct))
    if key == 'net_delivery_rate':
        return _band_caption(classify_band(value, net_delivery_target_pct, net_delivery_exceed_pct))
    if key == 'delivery_time_days' and _single_country in delivery_windows:
        return _band_caption(classify_delivery_time_band(value, delivery_windows[_single_country]))
    return None


if mode == "Single period":
    metrics = compute_period_metrics(df, start_date, end_date, countries=countries,
                                      on_time_target_days=on_time_target_days,
                                      net_delivery_matured_threshold=net_delivery_matured_threshold)

    st.download_button(
        "⬇️ Download this report as Excel",
        data=export_single_period_xlsx(metrics, export_meta),
        file_name=f"ops_pulse_{start_date}_to_{end_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tab_overview, = st.tabs(["Overview"])
    with tab_overview:
        st.subheader(f"{start_date} → {end_date}")
        render_metric_cards(metrics)

        st.subheader("Logistics KPIs by market -- vs. the CEO scorecard's bands")
        kpi_table = per_country_kpi_table(metrics)
        if not kpi_table.empty:
            st.dataframe(kpi_table, use_container_width=True, hide_index=True)
        else:
            st.caption("No markets in this selection have data for the period.")

        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(status_breakdown_chart(metrics, "Orders by status"), use_container_width=True)
        with col2:
            fig = per_country_rate_chart(metrics, 'delivered_rate', "Delivered rate by country")
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        col3, col4 = st.columns(2)
        with col3:
            fig = per_country_rate_chart(metrics, 'cancelled_rate', "Cancelled rate by country")
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with col4:
            fig = per_country_rate_chart(metrics, 'returned_rate', "Returned rate by country")
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        col5, col6 = st.columns(2)
        with col5:
            fig = per_country_rate_chart(metrics, 'on_time_rate', "On-time delivery rate by country")
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with col6:
            fig = per_country_rate_chart(metrics, 'net_delivery_rate', "Net delivery rate by country")
            if fig:
                st.plotly_chart(fig, use_container_width=True)

else:
    comparison = compare_periods(
        df, (a_start, a_end), (b_start, b_end), countries=countries,
        on_time_target_days=on_time_target_days,
        net_delivery_matured_threshold=net_delivery_matured_threshold,
    )
    metrics_a, metrics_b, deltas = comparison['period_a'], comparison['period_b'], comparison['deltas']
    summary = generate_summary(comparison, thresholds=thresholds)

    st.download_button(
        "⬇️ Download this report as Excel",
        data=export_comparison_xlsx(comparison, summary, export_meta),
        file_name=f"ops_pulse_{a_start}_to_{b_end}_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tab_overview, tab_comparison, tab_summary = st.tabs(["Overview", "Comparison", "Summary"])

    with tab_overview:
        st.subheader(f"Period B: {b_start} → {b_end}  (vs. Period A: {a_start} → {a_end})")
        render_metric_cards(metrics_b, delta_metrics=deltas)

        st.subheader("Logistics KPIs by market -- Period B, vs. the CEO scorecard's bands")
        kpi_table_b = per_country_kpi_table(metrics_b)
        if not kpi_table_b.empty:
            st.dataframe(kpi_table_b, use_container_width=True, hide_index=True)
        with st.expander("Same table for Period A (baseline)"):
            kpi_table_a = per_country_kpi_table(metrics_a)
            if not kpi_table_a.empty:
                st.dataframe(kpi_table_a, use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(status_breakdown_chart(metrics_a, f"Period A ({a_start} → {a_end}) -- by status"), use_container_width=True)
        with col2:
            st.plotly_chart(status_breakdown_chart(metrics_b, f"Period B ({b_start} → {b_end}) -- by status"), use_container_width=True)

    with tab_comparison:
        st.caption("Period A (baseline) vs. Period B, by market. Compare the % charts even when the two periods are different lengths.")
        for rate_key, label in [
            ('delivered_rate', 'Delivered rate'), ('cancelled_rate', 'Cancelled rate'),
            ('returned_rate', 'Returned rate'), ('pending_rate', 'Pending rate'),
            ('on_time_rate', 'On-time delivery rate'), ('net_delivery_rate', 'Net delivery rate'),
        ]:
            fig = comparison_bar_chart(comparison, rate_key, f"{label} by country -- A vs B")
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        st.subheader("Overall deltas (Period B − Period A)")
        st.caption(
            "Trend AND target status together: the Delta column says whether it moved "
            "in the right direction, the band columns say whether each period actually "
            "hit the CEO scorecard's target -- a metric can improve and still be Below "
            "target, or decline and still be On target."
        )
        delta_rows = []
        for key, label in METRIC_LABELS.items():
            d = deltas.get(key)
            if d is None:
                continue
            unit = 'pp' if key.endswith('_rate') else 'days'
            direction = METRIC_DIRECTION.get(key, 1)
            trend = ('↑ improving' if d * direction > 0 else '↓ declining') if d != 0 else '— flat'
            delta_rows.append({
                'Metric': label,
                'Period A': metrics_a.get(key), 'Period A band': _band_for_metric(key, metrics_a.get(key)) or '—',
                'Period B': metrics_b.get(key), 'Period B band': _band_for_metric(key, metrics_b.get(key)) or '—',
                'Delta': d, 'Unit': unit, 'Trend': trend,
            })
        if delta_rows:
            st.dataframe(pd.DataFrame(delta_rows), use_container_width=True, hide_index=True)
        if _single_country is None:
            st.caption(
                "Delivery Time's band needs one market at a time (its window differs by "
                "market) -- pick a single country in the sidebar to see it banded here, "
                "or check the per-market tables above/below."
            )

    with tab_summary:
        st.subheader("✅ What's working")
        if summary['good']:
            for item in summary['good']:
                st.success(item['message'])
        else:
            st.info("Nothing moved enough (given the current thresholds in the sidebar) to call it a clear improvement.")

        st.subheader("⚠️ Weak points")
        if summary['weak_points']:
            for item in summary['weak_points']:
                st.error(item['message'])
        else:
            st.success("No weak points crossed the thresholds in this comparison.")

        with st.expander("Thresholds used"):
            st.json(summary['thresholds_used'])
