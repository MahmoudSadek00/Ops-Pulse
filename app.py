import json
import datetime as dt

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from logic import (
    STATUSES, METRIC_LABELS, METRIC_DIRECTION, DEFAULT_ON_TIME_TARGET_DAYS,
    DEFAULT_WEAK_POINT_THRESHOLDS, STAGING_SPREADSHEET_ID_DEFAULT,
    get_client, load_orders_data, compute_period_metrics, compare_periods, generate_summary,
    export_single_period_xlsx, export_comparison_xlsx,
)

st.set_page_config(page_title="Ops Pulse", layout="wide")
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
            "On-time delivery target: a placeholder (5 days) until the CEO scorecard's "
            "real per-market transit windows are plugged in here -- applies to every "
            "market equally for now."
        )
        on_time_target = st.number_input(
            "On-time delivery target (days from Order Date)",
            min_value=1, max_value=60, value=DEFAULT_ON_TIME_TARGET_DAYS,
        )
        st.caption("Weak-point thresholds -- how much worse than the baseline period counts as a weak point.")
        thresholds = {}
        thresholds['delivered_rate_pp'] = st.number_input("Delivered rate drop (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['delivered_rate_pp'], 0.5)
        thresholds['cancelled_rate_pp'] = st.number_input("Cancelled rate rise (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['cancelled_rate_pp'], 0.5)
        thresholds['returned_rate_pp'] = st.number_input("Returned rate rise (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['returned_rate_pp'], 0.5)
        thresholds['pending_rate_pp'] = st.number_input("Pending rate rise (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['pending_rate_pp'], 0.5)
        thresholds['fulfillment_lead_time_days'] = st.number_input("Fulfillment lead time slowdown (days)", 0.1, 10.0, DEFAULT_WEAK_POINT_THRESHOLDS['fulfillment_lead_time_days'], 0.1)
        thresholds['delivery_time_days'] = st.number_input("Delivery time slowdown (days)", 0.1, 10.0, DEFAULT_WEAK_POINT_THRESHOLDS['delivery_time_days'], 0.1)
        thresholds['on_time_rate_pp'] = st.number_input("On-time rate drop (pp)", 0.5, 20.0, DEFAULT_WEAK_POINT_THRESHOLDS['on_time_rate_pp'], 0.5)

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
    return f"{x:,.0f}" if x is not None else "—"


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
               help=f"n={metrics['delivery_time_n']:,} Delivered orders with both an Order Date and a Delivery Date.")
    c8.metric("On-time delivery rate", _pct(metrics['on_time_rate']),
               delta=(f"{delta_metrics['on_time_rate']:+.1f}pp" if delta_metrics and delta_metrics.get('on_time_rate') is not None else None),
               delta_color=_dc('on_time_rate'),
               help=f"Against a {on_time_target}-day target. n={metrics['on_time_n']:,}.")

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
    'on_time_target_days': on_time_target,
    'generated_at': dt.datetime.now().strftime('%Y-%m-%d %H:%M'),
}

if mode == "Single period":
    metrics = compute_period_metrics(df, start_date, end_date, countries=countries, on_time_target_days=on_time_target)

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

else:
    comparison = compare_periods(
        df, (a_start, a_end), (b_start, b_end), countries=countries, on_time_target_days=on_time_target,
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
        ]:
            fig = comparison_bar_chart(comparison, rate_key, f"{label} by country -- A vs B")
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        st.subheader("Overall deltas (Period B − Period A)")
        delta_rows = []
        for key, label in METRIC_LABELS.items():
            d = deltas.get(key)
            if d is None:
                continue
            unit = 'pp' if key.endswith('_rate') else 'days'
            delta_rows.append({'Metric': label, 'Period A': metrics_a.get(key), 'Period B': metrics_b.get(key),
                                'Delta': d, 'Unit': unit})
        if delta_rows:
            st.dataframe(pd.DataFrame(delta_rows), use_container_width=True, hide_index=True)

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
