import asyncio
import html
import io
import logging
import re
from typing import Any, Dict, List, Optional, Tuple

import altair as alt
import httpx
import pandas as pd
import streamlit as st

from genie_client_streamlit import GenieClient, GenieError
from streamlit_runtime import settings
from streamlit_store import store


# ============================================================================
# Page configuration
# ============================================================================

st.set_page_config(
    page_title="TNS Retail Intelligence",
    page_icon="https://admin.thenewshop.in/static/media/New%20Logo%20.ad69756dd0621a9db47a.jpg",
    layout="wide",
    initial_sidebar_state="expanded",
)

LOGO_URL = "https://admin.thenewshop.in/static/media/New%20Logo%20.ad69756dd0621a9db47a.jpg"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("tns_streamlit")


# ============================================================================
# CSS
# ============================================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        #MainMenu {visibility:hidden;}
        footer {visibility:hidden;}
        header {background:transparent !important;}
        .stApp {background:#f5f7fb; color:#172033;}
        [data-testid="stSidebar"] {background:#111827; min-width:270px; max-width:270px;}
        [data-testid="stSidebar"] * {color:#d1d5db;}
        [data-testid="stSidebar"] .stButton button {
            text-align:left;
            border:1px solid #374151;
            background:#1f2937;
            color:white;
            border-radius:9px;
        }
        [data-testid="stSidebar"] .stButton button:hover {
            background:#374151;
            border-color:#4b5563;
        }
        .tns-brand {display:flex;align-items:center;gap:10px;padding:4px 0 18px;}
        .tns-brand img {width:38px;height:38px;border-radius:9px;object-fit:contain;background:white;}
        .tns-brand-name {font-weight:650;font-size:16px;color:white;}
        .tns-header {
            background:white;
            border-bottom:1px solid #e5e7eb;
            padding:12px 24px;
            margin:-1rem -1rem 1rem;
        }
        .tns-header-title {font-weight:600;font-size:15px;color:#172033;}
        .tns-header-subtitle {font-size:11px;color:#9ca3af;margin-top:2px;}
        .welcome {text-align:center;margin:12vh auto 8vh;}
        .welcome h1 {font-size:30px;letter-spacing:-.5px;color:#172033;margin-bottom:10px;}
        .welcome p {font-size:14px;color:#6b7280;}
        .user-bubble {
            background:#111827;
            color:white;
            border-radius:14px 14px 4px 14px;
            padding:13px 16px;
            line-height:1.55;
            margin:12px 0 18px 18%;
        }
        .assistant-bubble {
            background:white;
            border:1px solid #e5e7eb;
            border-radius:14px 14px 14px 4px;
            padding:13px 16px;
            line-height:1.6;
            margin:12px 18% 18px 0;
        }
        .section-label {
            font-size:11px;
            font-weight:650;
            text-transform:uppercase;
            letter-spacing:.7px;
            color:#9ca3af;
            margin:14px 0 7px;
        }
        .result-card {
            border:1px solid #e5e7eb;
            border-radius:12px;
            background:white;
            padding:14px;
            margin:12px 0;
        }
        .result-title {font-weight:650;color:#172033;margin-bottom:8px;}
        .source-note {font-size:11px;color:#9ca3af;margin-top:7px;}
        .chart-title {font-size:17px;font-weight:650;color:#172033;margin-bottom:2px;}
        .chart-description {font-size:12px;color:#6b7280;margin-bottom:8px;}
        .chart-source {font-size:11px;color:#9ca3af;margin-top:4px;}
        .login-wrap {
            max-width:430px;
            margin:12vh auto 0;
            background:white;
            border:1px solid #e5e7eb;
            border-radius:16px;
            padding:34px;
            box-shadow:0 10px 30px rgba(17,24,39,.06);
        }
        .login-logo {
            display:block;
            width:58px;
            height:58px;
            object-fit:contain;
            margin:0 auto 15px;
            border-radius:12px;
            background:white;
        }
        .login-title{text-align:center;font-size:26px;font-weight:700;color:#172033;}
        .login-sub{text-align:center;color:#6b7280;font-size:13px;margin:7px 0 24px;}
        div[data-testid="stChatMessage"] {background:transparent;}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# Async / Genie helpers
# ============================================================================

def run_async(coro):
    return asyncio.run(coro)


def new_client() -> GenieClient:
    return GenieClient()


def call_genie(coro_factory):
    client = new_client()
    try:
        return run_async(coro_factory(client))
    finally:
        # Keep the same lifecycle used by the working Streamlit app.
        try:
            run_async(client.close())
        except RuntimeError as exc:
            # Do not turn an already completed Genie response into a UI error
            # merely because httpx is trying to close an event-loop-bound
            # transport after asyncio.run() has finished.
            logger.warning("Genie client close skipped: %s", exc)


# ============================================================================
# Authentication / storage
# ============================================================================

def check_credentials(username: str, password: str) -> bool:
    return username == settings.app_username and password == settings.app_password


def ensure_store() -> None:
    if st.session_state.get("store_initialized"):
        return
    store.initialize()
    st.session_state.store_initialized = True


def login_screen() -> None:
    inject_css()
    st.markdown(
        f'''<div class="login-wrap">
            <img class="login-logo" src="{LOGO_URL}" />
            <div class="login-title">TNS Retail Intelligence</div>
            <div class="login-sub">Sign in to access company analytics</div>
        </div>''',
        unsafe_allow_html=True,
    )

    with st.form("login_form"):
        username = st.text_input("Username", autocomplete="username")
        password = st.text_input(
            "Password",
            type="password",
            autocomplete="current-password",
        )
        submitted = st.form_submit_button("Sign in", use_container_width=True)

    if submitted:
        if check_credentials(username.strip(), password):
            st.session_state.authenticated = True
            st.session_state.username = username.strip()
            st.session_state.active_chat_id = None
            st.rerun()
        else:
            st.error("Invalid username or password.")


def get_sessions():
    return store.list_sessions(st.session_state.username)


def load_history(session_id: str):
    history = store.get_history(session_id, st.session_state.username)
    if history is None:
        st.error("Chat session not found.")
        return None
    return history


# ============================================================================
# Query-result rendering
# ============================================================================

def download_signed_links(
    table: Dict[str, Any],
    conversation_id: str,
    message_id: str,
) -> None:
    attachment_id = table.get("attachment_id")
    if not attachment_id or not message_id:
        return

    try:
        links = call_genie(
            lambda client: client.get_full_query_download_links(
                conversation_id,
                message_id,
                attachment_id,
            )
        )
        for index, link in enumerate(links):
            try:
                response = httpx.get(link, timeout=120.0)
                response.raise_for_status()
                label = (
                    "Download query result"
                    if len(links) == 1
                    else f"Download query result {index + 1}"
                )
                st.download_button(
                    label,
                    data=response.content,
                    file_name=f"genie_query_result_{index + 1}.csv",
                    mime="text/csv",
                    key=f"download_{conversation_id}_{message_id}_{attachment_id}_{index}",
                )
            except Exception as exc:
                logger.warning("Could not download query result: %s", exc)
                st.link_button(f"Open download {index + 1}", link)
    except Exception as exc:
        logger.warning("Could not create query-result download: %s", exc)


def render_table(
    table: Dict[str, Any],
    conversation_id: str,
    message_id: str,
) -> None:
    with st.container(border=True):
        st.markdown(f"**{table.get('title') or 'Query result'}**")
        if table.get("description"):
            st.caption(table["description"])

        columns = table.get("columns") or []
        rows = table.get("rows") or []

        if columns:
            df = pd.DataFrame(rows, columns=columns)
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info(table.get("error") or "No tabular result was returned.")

        displayed = table.get("displayed_row_count", len(rows))
        total = table.get("row_count", displayed)
        if table.get("truncated") or total > displayed:
            st.caption(f"Showing {displayed:,} of {total:,} rows.")

        download_signed_links(table, conversation_id, message_id)


# ============================================================================
# Semantic chart engine
#
# IMPORTANT:
#   This does NOT download the Databricks visualization PNG.
#   It uses the exact query attachment referenced by
#   viz.query_attachment_id and renders that real Genie result with
#   a deterministic Vega-Lite/Altair specification.
# ============================================================================

_CURRENCY_WORDS = (
    "revenue", "sales", "mrp", "amount", "price", "cost", "profit",
    "salesvalue", "salesamount", "gmv", "turnover", "discount",
    "deliveryfee", "taxamount", "netamount", "grossamount", "spend",
    "expense", "expenditure", "value",
)

_QUANTITY_WORDS = (
    "quantity", "qty", "unit", "units", "unitcount", "itemcount",
    "productcount", "skucount", "ordercount", "storecount",
    "customercount", "count", "volume", "demand", "orders", "transactions",
    "invoices", "bills",
)

_PERCENT_WORDS = (
    "percent", "percentage", "pct", "rate", "margin", "share", "growth",
    "contribution",
)

_SERIES_GROUPS = {
    "quantity": (
        "quantity", "qty", "units", "unit", "demand", "volume",
    ),
    "revenue": (
        "revenue", "sales", "amount", "value", "gmv", "turnover",
    ),
    "orders": (
        "orders", "order", "transactions", "transaction", "invoices",
        "invoice", "bills", "bill", "count",
    ),
    "cost": ("cost", "cogs", "purchase", "expense", "expenditure"),
    "profit": ("profit", "grossprofit", "netprofit"),
    "margin": ("margin", "rate", "percentage", "percent", "pct"),
    "forecast": ("forecast", "predicted", "prediction", "projected", "estimate"),
}


def normalize_column_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def is_numeric_value(value: Any) -> bool:
    if value is None or value == "":
        return False
    try:
        if isinstance(value, bool):
            return False
        number = float(str(value).strip())
        return pd.notna(number)
    except (TypeError, ValueError):
        return False


def is_iso_date(value: Any) -> bool:
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}([T ]|$)", str(value)))


def is_currency_column(column_name: Any) -> bool:
    name = normalize_column_name(column_name)
    if is_quantity_column(column_name) or is_percentage_column(column_name):
        return False
    return any(word in name for word in _CURRENCY_WORDS)


def is_quantity_column(column_name: Any) -> bool:
    name = normalize_column_name(column_name)
    return (
        name == "quantity"
        or "quantity" in name
        or name == "qty"
        or "qty" in name
        or "units" in name
        or "unitcount" in name
        or "itemcount" in name
        or "productcount" in name
        or "skucount" in name
        or "ordercount" in name
        or "storecount" in name
        or "customercount" in name
        or name.endswith("count")
        or "demand" in name
        or "volume" in name
    )


def is_percentage_column(column_name: Any) -> bool:
    name = normalize_column_name(column_name)
    return (
        "percent" in name
        or "percentage" in name
        or name.endswith("pct")
        or "rate" in name
        or "margin" in name
        or "share" in name
        or "growth" in name
        or "contribution" in name
    )


def column_kind(column_name: Any) -> str:
    if is_percentage_column(column_name):
        return "percent"
    if is_currency_column(column_name):
        return "currency"
    if is_quantity_column(column_name):
        return "quantity"
    return "number"


def format_indian_currency(value: Any) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)

    if not pd.notna(number):
        return str(value)

    absolute = abs(number)
    sign = "-" if number < 0 else ""
    if absolute >= 10_000_000:
        return f"{sign}₹{absolute / 10_000_000:,.2f} Cr"
    if absolute >= 100_000:
        return f"{sign}₹{absolute / 100_000:,.2f} L"
    return f"{sign}₹{absolute:,.2f}"


def format_chart_value(value: Any, kind: str) -> str:
    if value is None:
        return "—"
    if kind == "currency":
        return format_indian_currency(value)
    try:
        number = float(value)
        text = f"{number:,.2f}".rstrip("0").rstrip(".")
    except (TypeError, ValueError):
        return str(value)
    return f"{text}%" if kind == "percent" else text


def format_axis_label(value: Any, monthly: bool = False) -> str:
    if not is_iso_date(value):
        return str(value)
    try:
        date = pd.to_datetime(str(value), utc=True)
        if monthly:
            return date.strftime("%b %Y")
        return date.strftime("%d %b %Y")
    except Exception:
        return str(value)


def title_tokens(text: str) -> List[str]:
    return re.findall(r"[a-z0-9]+", text.lower())


def column_matches_group(column: str, group: str) -> bool:
    normalized = normalize_column_name(column)
    words = _SERIES_GROUPS.get(group, ())
    return any(word in normalized for word in words)


def title_mentions_group(title: str, group: str) -> bool:
    normalized = normalize_column_name(title)
    words = _SERIES_GROUPS.get(group, ())
    return any(word in normalized for word in words)


def infer_requested_groups(viz: Dict[str, Any]) -> List[str]:
    text = " ".join(
        str(viz.get(key) or "")
        for key in ("title", "description", "chart_type", "type")
    ).lower()

    # More specific semantic groups first.
    ordered = [
        "margin",
        "profit",
        "forecast",
        "cost",
        "revenue",
        "quantity",
        "orders",
    ]

    groups = [group for group in ordered if title_mentions_group(text, group)]

    # "profit margin" should be a margin chart, not merely profit.
    if "margin" in groups:
        groups = ["margin"] + [g for g in groups if g != "margin"]

    return groups


def infer_chart_type(viz: Dict[str, Any], df: pd.DataFrame) -> str:
    hint = " ".join(
        str(viz.get(key) or "")
        for key in ("chart_type", "type", "title", "description")
    ).lower()

    if "pie" in hint or "donut" in hint:
        return "pie"

    first_col_is_date = bool(len(df) and df.iloc[:, 0].map(is_iso_date).all())

    if first_col_is_date or any(word in hint for word in ("line", "trend", "over time", "monthly", "weekly", "daily")):
        return "line"

    return "bar"


def find_source_table(viz: Dict[str, Any], tables: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Strictly resolve the visualization to its query attachment.

    We intentionally DO NOT fall back to an arbitrary table. That was one of
    the causes of meaningless charts: a visualization could accidentally use
    the last numerical table in the message instead of the query that Genie
    associated with that visualization.
    """
    query_attachment_id = viz.get("query_attachment_id")
    if not query_attachment_id:
        return None

    normalized_target = str(query_attachment_id)
    for table in tables:
        if str(table.get("attachment_id") or "") == normalized_target:
            return table
    return None


def choose_measure_columns(
    df: pd.DataFrame,
    viz: Dict[str, Any],
) -> Tuple[List[str], str]:
    numeric_columns: List[str] = []
    for column in df.columns[1:]:
        values = df[column].dropna()
        if not values.empty and values.map(is_numeric_value).all():
            numeric_columns.append(str(column))

    if not numeric_columns:
        return [], "number"

    requested_groups = infer_requested_groups(viz)

    preferred: List[str] = []
    for group in requested_groups:
        for column in numeric_columns:
            if column_matches_group(column, group) and column not in preferred:
                preferred.append(column)

    # If the visualization title does not identify the measure, use the first
    # numeric column's semantic kind and keep related columns such as actual /
    # predicted / lower / upper bounds together.
    if not preferred:
        first_kind = column_kind(numeric_columns[0])
        preferred = [
            column for column in numeric_columns
            if column_kind(column) == first_kind
        ]

    # Forecasts are special: if the result contains actual + predicted + bounds,
    # show those together rather than only one arbitrary numerical column.
    forecast_like = [
        column for column in numeric_columns
        if any(
            token in normalize_column_name(column)
            for token in ("actual", "forecast", "predicted", "prediction", "lower", "upper", "bound")
        )
    ]
    if "forecast" in requested_groups and forecast_like:
        preferred = forecast_like

    preferred = preferred[:5]
    first_kind = column_kind(preferred[0]) if preferred else column_kind(numeric_columns[0])

    # Add compatible secondary series only when they are semantically related.
    selected = list(preferred)
    for column in numeric_columns:
        if column in selected:
            continue
        if len(selected) >= 5:
            break
        if column_kind(column) == first_kind:
            normalized = normalize_column_name(column)
            if any(
                token in normalized
                for token in ("actual", "forecast", "predicted", "prediction", "target", "budget", "plan")
            ):
                selected.append(column)

    return selected, first_kind


def build_chart_dataframe(
    table: Dict[str, Any],
    viz: Dict[str, Any],
) -> Optional[Tuple[pd.DataFrame, List[str], str, str, bool]]:
    columns = table.get("columns") or []
    rows = table.get("rows") or []
    if len(columns) < 2 or not rows:
        return None

    # Keep the complete displayed Genie result. Do not randomly sample it.
    df = pd.DataFrame(rows, columns=columns)

    # Normalize numeric values only in selected measure columns.
    measures, first_kind = choose_measure_columns(df, viz)
    if not measures:
        return None

    for column in measures:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    x_column = str(df.columns[0])
    date_axis = bool(df[x_column].map(is_iso_date).all())
    monthly = date_axis and bool(
        df[x_column].astype(str).str.match(r"^\d{4}-\d{2}-01").all()
    )

    # Friendly display labels, while retaining the original source values.
    display_x = f"__display_{x_column}"
    df[display_x] = df[x_column].map(lambda value: format_axis_label(value, monthly))

    chart_type = infer_chart_type(viz, df)

    # Pie charts only make sense for a small part-to-whole result.
    if chart_type == "pie":
        if len(df) > 8:
            chart_type = "bar"
        elif len(measures) != 1:
            chart_type = "bar"

    return df, measures, chart_type, first_kind, date_axis


def build_altair_chart(
    table: Dict[str, Any],
    viz: Dict[str, Any],
) -> Optional[Tuple[alt.Chart, pd.DataFrame, str]]:
    built = build_chart_dataframe(table, viz)
    if not built:
        return None

    df, measures, chart_type, first_kind, date_axis = built
    x_column = str(df.columns[0])
    display_x = f"__display_{x_column}"

    title = str(viz.get("title") or "Genie visualization")
    description = str(viz.get("description") or "")

    tooltip_fields = [
        alt.Tooltip(display_x + ":N", title=x_column),
    ]
    for measure in measures:
        tooltip_fields.append(
            alt.Tooltip(
                f"{measure}:Q",
                title=measure,
                format=",.2f",
            )
        )

    base = alt.Chart(df).properties(
        title=alt.TitleParams(
            text=title,
            subtitle=description if description else None,
            anchor="start",
            fontSize=17,
            subtitleFontSize=12,
        ),
        height=max(340, min(650, 80 + len(df) * 24)) if chart_type == "bar" and len(df) > 12 else 360,
    )

    if chart_type == "pie":
        measure = measures[0]
        chart = (
            base.mark_arc(innerRadius=55)
            .encode(
                theta=alt.Theta(field=measure, type="quantitative"),
                color=alt.Color(
                    field=display_x,
                    type="nominal",
                    title=x_column,
                    legend=alt.Legend(orient="right"),
                ),
                tooltip=tooltip_fields,
            )
            .interactive()
        )
        return chart, df, first_kind

    if chart_type == "line":
        layers = []
        for measure in measures:
            line = (
                alt.Chart(df)
                .mark_line(point=True, strokeWidth=2.5)
                .encode(
                    x=alt.X(
                        f"{display_x}:N",
                        title=x_column,
                        sort=None,
                        axis=alt.Axis(labelAngle=-35),
                    ),
                    y=alt.Y(
                        f"{measure}:Q",
                        title=format_axis_title(measure, first_kind),
                        axis=alt.Axis(format=axis_format(first_kind)),
                    ),
                    color=alt.value(chart_color(len(layers))),
                    tooltip=tooltip_fields,
                )
            )
            layers.append(line)

        chart = alt.layer(*layers).resolve_scale(y="shared").properties(
            title=alt.TitleParams(
                text=title,
                subtitle=description if description else None,
                anchor="start",
                fontSize=17,
                subtitleFontSize=12,
            ),
            height=360,
        ).interactive()
        return chart, df, first_kind

    # Bar chart. For many categories, use a horizontal layout for readability.
    horizontal = len(df) > 12
    if horizontal:
        y = alt.Y(
            f"{display_x}:N",
            title=x_column,
            sort="-x",
            axis=alt.Axis(labelLimit=260),
        )
        x = alt.X(
            f"{measures[0]}:Q",
            title=format_axis_title(measures[0], first_kind),
            axis=alt.Axis(format=axis_format(first_kind)),
        )
        layers = []
        for index, measure in enumerate(measures):
            bar = (
                alt.Chart(df)
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=y,
                    x=alt.X(
                        f"{measure}:Q",
                        title=format_axis_title(measure, column_kind(measure)),
                        axis=alt.Axis(format=axis_format(column_kind(measure))),
                    ),
                    color=alt.value(chart_color(index)),
                    tooltip=tooltip_fields,
                )
            )
            layers.append(bar)
        chart = alt.layer(*layers).resolve_scale(x="independent")
    else:
        x = alt.X(
            f"{display_x}:N",
            title=x_column,
            sort=None,
            axis=alt.Axis(labelAngle=-35, labelLimit=180),
        )
        layers = []
        for index, measure in enumerate(measures):
            bar = (
                alt.Chart(df)
                .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=x,
                    y=alt.Y(
                        f"{measure}:Q",
                        title=format_axis_title(measure, column_kind(measure)),
                        axis=alt.Axis(format=axis_format(column_kind(measure))),
                    ),
                    color=alt.value(chart_color(index)),
                    tooltip=tooltip_fields,
                )
            )
            layers.append(bar)
        chart = alt.layer(*layers).resolve_scale(y="independent")

    chart = chart.properties(
        title=alt.TitleParams(
            text=title,
            subtitle=description if description else None,
            anchor="start",
            fontSize=17,
            subtitleFontSize=12,
        ),
        height=max(340, min(650, 80 + len(df) * 24)) if horizontal else 360,
    ).interactive()

    return chart, df, first_kind


def chart_color(index: int) -> str:
    # Deliberately controlled, restrained palette. These are chart-rendering
    # colors, not data values; the data itself always comes from Genie.
    colors = ["#2563eb", "#f59e0b", "#10b981", "#ef4444", "#8b5cf6"]
    return colors[index % len(colors)]


def axis_format(kind: str) -> str:
    if kind == "currency":
        return ",.2s"
    if kind == "percent":
        return ".1f"
    return ",.2f"


def format_axis_title(column: str, kind: str) -> str:
    if kind == "currency":
        return f"{column} (₹)"
    if kind == "percent":
        return f"{column} (%)"
    return column


def chart_to_png(df: pd.DataFrame, measures: List[str], chart_type: str, title: str, first_kind: str) -> bytes:
    """Create a deterministic downloadable PNG from the SAME chart data.

    The on-screen chart is Altair/Vega-Lite. This fallback export uses
    matplotlib only for downloading because Streamlit/Altair PNG export may
    require an additional browser renderer package. It never changes the
    on-screen chart or selects different data.
    """
    import matplotlib.pyplot as plt

    fig = plt.figure(figsize=(12, 6.5))
    ax = fig.add_subplot(111)

    x = df.iloc[:, 0].astype(str).tolist()
    display_x = [format_axis_label(value, False) for value in x]

    if chart_type == "pie" and measures:
        values = pd.to_numeric(df[measures[0]], errors="coerce").fillna(0)
        ax.pie(values, labels=display_x, autopct="%1.1f%%")
    elif chart_type == "line":
        for measure in measures:
            ax.plot(display_x, pd.to_numeric(df[measure], errors="coerce"), marker="o", label=measure)
        if len(measures) > 1:
            ax.legend()
        ax.tick_params(axis="x", rotation=35)
    else:
        if len(measures) == 1:
            values = pd.to_numeric(df[measures[0]], errors="coerce").fillna(0)
            if len(df) > 12:
                ax.barh(display_x, values)
                ax.set_xlabel(format_axis_title(measures[0], first_kind))
            else:
                ax.bar(display_x, values)
                ax.tick_params(axis="x", rotation=35)
                ax.set_ylabel(format_axis_title(measures[0], first_kind))
        else:
            width = 0.8 / len(measures)
            positions = list(range(len(df)))
            for index, measure in enumerate(measures):
                values = pd.to_numeric(df[measure], errors="coerce").fillna(0)
                offset = [p - 0.4 + width / 2 + index * width for p in positions]
                ax.bar(offset, values, width=width, label=measure)
            ax.set_xticks(positions)
            ax.set_xticklabels(display_x, rotation=35, ha="right")
            ax.legend()

    ax.set_title(title, loc="left", pad=15, fontweight="bold")
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    return buffer.getvalue()


def render_visualization(
    viz: Dict[str, Any],
    tables: List[Dict[str, Any]],
    conversation_id: str,
    message_id: str,
) -> Optional[str]:
    """Render a Genie visualization from its actual query result.

    Returns the source query attachment ID when successfully rendered, so the
    caller can avoid rendering the same table again below the chart.
    """
    title = viz.get("title") or "Genie visualization"
    table = find_source_table(viz, tables)

    with st.container(border=True):
        st.markdown(f'<div class="chart-title">{html.escape(str(title))}</div>', unsafe_allow_html=True)
        if viz.get("description"):
            st.markdown(
                f'<div class="chart-description">{html.escape(str(viz["description"]))}</div>',
                unsafe_allow_html=True,
            )

        if table is None:
            st.warning(
                "Genie returned a visualization without a matching query attachment, "
                "so no chart was fabricated from another table."
            )
            return None

        try:
            built = build_chart_dataframe(table, viz)
            if not built:
                st.warning("The Genie query result does not contain a chartable dimension and measure.")
                return None

            df, measures, chart_type, first_kind, _ = built
            chart_result = build_altair_chart(table, viz)
            if not chart_result:
                st.warning("The Genie query result could not be converted into a chart.")
                return None

            chart, _, _ = chart_result
            st.altair_chart(chart, use_container_width=True)

            st.markdown(
                f'<div class="chart-source">Source: {html.escape(str(table.get("title") or "Query result"))}</div>',
                unsafe_allow_html=True,
            )

            # Keep the actual data available, just as the original FastAPI UI
            # exposed "View data" below a native chart.
            with st.expander("View data", expanded=False):
                source_df = pd.DataFrame(
                    table.get("rows") or [],
                    columns=table.get("columns") or [],
                )
                st.dataframe(source_df, use_container_width=True, hide_index=True)

            png = chart_to_png(
                df=df,
                measures=measures,
                chart_type=chart_type,
                title=str(title),
                first_kind=first_kind,
            )
            st.download_button(
                "Download chart",
                data=png,
                file_name="genie_visualization.png",
                mime="image/png",
                key=f"viz_download_{conversation_id}_{message_id}_{viz.get('attachment_id') or title}",
            )

            return str(table.get("attachment_id") or "")
        except Exception as exc:
            logger.exception("Semantic Genie chart rendering failed")
            st.error(f"Could not render the Genie chart: {exc}")
            return None


# ============================================================================
# Presentation rendering
# ============================================================================

def render_thoughts(thoughts):
    if not thoughts:
        return
    with st.expander("Thought process", expanded=False):
        for thought in thoughts:
            content = thought.get("content") if isinstance(thought, dict) else str(thought)
            if content:
                st.markdown(content)


def render_presentation(
    presentation: Dict[str, Any],
    conversation_id: str,
    message_id: str,
) -> None:
    if not presentation:
        return

    thoughts = presentation.get("thoughts") or []
    if not thoughts:
        for block in presentation.get("blocks") or []:
            if block.get("type") == "thoughts":
                thoughts = block.get("data") or []
                break
    render_thoughts(thoughts)

    visualizations = presentation.get("visualizations") or []
    if not visualizations:
        visualizations = [
            b.get("data")
            for b in presentation.get("blocks") or []
            if b.get("type") == "visualization"
        ]
    visualizations = [v for v in visualizations if isinstance(v, dict)]

    tables = presentation.get("tables") or []
    if not tables:
        tables = [
            b.get("data")
            for b in presentation.get("blocks") or []
            if b.get("type") == "table"
        ]
    tables = [t for t in tables if isinstance(t, dict)]

    chart_source_ids = set()
    for viz in visualizations:
        source_id = render_visualization(
            viz,
            tables,
            conversation_id,
            message_id,
        )
        if source_id:
            chart_source_ids.add(source_id)

    # Only hide a table when it is EXACTLY the query attachment used by a
    # successfully rendered visualization. Never hide a table based on a
    # guessed/fallback association.
    for table in tables:
        attachment_id = str(table.get("attachment_id") or "")
        if attachment_id in chart_source_ids:
            continue
        render_table(table, conversation_id, message_id)

    suggestions = presentation.get("suggested_questions") or []
    if not suggestions:
        for block in presentation.get("blocks") or []:
            if block.get("type") == "suggested_questions":
                suggestions = block.get("data") or []
                break

    if suggestions:
        st.markdown(
            '<div class="section-label">Suggested questions</div>',
            unsafe_allow_html=True,
        )
        for index, question in enumerate(suggestions):
            if st.button(
                question,
                key=f"suggestion_{message_id}_{index}",
                use_container_width=True,
            ):
                st.session_state.pending_prompt = question
                st.rerun()


def render_message(message: Dict[str, Any], conversation_id: str) -> None:
    role = message.get("role")
    content = message.get("content") or ""

    if role == "user":
        safe = html.escape(content).replace("\n", "<br>")
        st.markdown(f'<div class="user-bubble">{safe}</div>', unsafe_allow_html=True)
        return

    st.markdown("<div class='assistant-bubble'>", unsafe_allow_html=True)
    st.markdown(content)
    st.markdown("</div>", unsafe_allow_html=True)

    presentation = message.get("presentation") or message.get("metadata") or {}
    message_id = (
        presentation.get("agent_message_id")
        or message.get("message_id")
        or ""
    )
    render_presentation(presentation, conversation_id, message_id)


# ============================================================================
# Agent turn processing
# ============================================================================

async def run_agent_turn(
    client: GenieClient,
    message: str,
    conversation_id: Optional[str],
):
    original = conversation_id
    rebound = False

    try:
        response = await client.create_agent_response(
            message,
            conversation_id=conversation_id,
            enable_visualization=True,
        )
    except GenieError as exc:
        if conversation_id and client.is_legacy_conversation_error(exc):
            rebound = True
            response = await client.create_agent_response(
                message,
                conversation_id=None,
                enable_visualization=True,
            )
        else:
            raise

    resolved = response.get("conversation_id") or (None if rebound else conversation_id)
    if not resolved:
        raise GenieError("Genie Agent did not return a conversation ID.")

    answer = client.normalize_answer_text(client.extract_agent_answer(response))
    presentation = await client.build_agent_presentation(response)
    message_id = presentation.get("agent_message_id") or response.get("id") or ""

    return {
        "answer": answer,
        "presentation": presentation,
        "conversation_id": resolved,
        "conversation_changed": resolved != original,
        "message_id": message_id,
    }


def process_new_message(prompt: str):
    result = call_genie(lambda client: run_agent_turn(client, prompt, None))
    title = prompt.strip()
    if len(title) > 60:
        title = title[:57] + "..."

    sid = store.create_session(
        result["conversation_id"],
        st.session_state.username,
        title,
    )
    store.add_message(sid, st.session_state.username, "user", prompt)
    store.add_message(
        sid,
        st.session_state.username,
        "assistant",
        result["answer"],
        result["presentation"],
    )
    st.session_state.active_chat_id = sid
    return result


def process_followup(session_id: str, prompt: str):
    conversation_id = store.get_conversation_id(
        session_id,
        st.session_state.username,
    )
    if not conversation_id:
        raise GenieError("Chat session not found.")

    result = call_genie(
        lambda client: run_agent_turn(client, prompt, conversation_id)
    )

    if result["conversation_changed"]:
        store.set_conversation_id(
            session_id,
            st.session_state.username,
            result["conversation_id"],
        )

    store.add_message(session_id, st.session_state.username, "user", prompt)
    store.add_message(
        session_id,
        st.session_state.username,
        "assistant",
        result["answer"],
        result["presentation"],
    )
    return result


# ============================================================================
# Sidebar
# ============================================================================

def render_sidebar() -> None:
    with st.sidebar:
        st.markdown(
            f'<div class="tns-brand"><img src="{LOGO_URL}"/><div class="tns-brand-name">TNS Retail Intelligence</div></div>',
            unsafe_allow_html=True,
        )

        if st.button("＋  New Chat", use_container_width=True):
            st.session_state.active_chat_id = None
            st.session_state.pending_prompt = ""
            st.rerun()

        st.markdown(
            '<div class="section-label">Conversations</div>',
            unsafe_allow_html=True,
        )

        sessions = get_sessions()
        for session in sessions:
            label = session["title"] or "New Chat"
            c1, c2 = st.columns([0.86, 0.14], gap="small")
            with c1:
                if st.button(
                    label,
                    key=f"chat_{session['session_id']}",
                    use_container_width=True,
                ):
                    st.session_state.active_chat_id = session["session_id"]
                    st.session_state.pending_prompt = ""
                    st.rerun()
            with c2:
                if st.button(
                    "×",
                    key=f"delete_{session['session_id']}",
                    help="Delete conversation",
                ):
                    store.delete_session(
                        session["session_id"],
                        st.session_state.username,
                    )
                    if st.session_state.get("active_chat_id") == session["session_id"]:
                        st.session_state.active_chat_id = None
                    st.rerun()

        st.divider()
        st.caption(st.session_state.username)

        if st.button("Logout", use_container_width=True):
            st.session_state.clear()
            st.rerun()


# ============================================================================
# Main application
# ============================================================================

def main_app() -> None:
    inject_css()
    ensure_store()
    render_sidebar()

    active_id = st.session_state.get("active_chat_id")
    history = load_history(active_id) if active_id else None
    title = history["title"] if history else "New Chat"

    st.markdown(
        f'''<div class="tns-header">
            <div class="tns-header-title">{html.escape(str(title))}</div>
            <div class="tns-header-subtitle">Business Intelligence Assistant</div>
        </div>''',
        unsafe_allow_html=True,
    )

    if history:
        conversation_id = store.get_conversation_id(
            active_id,
            st.session_state.username,
        )
        for message in history["messages"]:
            render_message(message, conversation_id)
    else:
        st.markdown(
            '''<div class="welcome">
                <h1>How can I help you?</h1>
                <p>Ask questions about company sales, products, stores, brands and more.</p>
            </div>''',
            unsafe_allow_html=True,
        )

    pending = st.session_state.pop("pending_prompt", "")
    prompt = st.chat_input("Ask a question...", max_chars=5000)
    if pending:
        prompt = pending

    if prompt:
        with st.spinner("Genie is analyzing your request..."):
            try:
                if active_id:
                    process_followup(active_id, prompt)
                else:
                    process_new_message(prompt)
                st.rerun()
            except Exception as exc:
                logger.exception("Genie request failed")
                st.error(f"Request failed: {exc}")


# ============================================================================
# Entrypoint
# ============================================================================

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "pending_prompt" not in st.session_state:
    st.session_state.pending_prompt = ""

if st.session_state.authenticated:
    try:
        main_app()
    except Exception as exc:
        logger.exception("Application error")
        st.error(f"Application error: {exc}")
else:
    login_screen()
