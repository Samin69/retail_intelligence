import asyncio
import csv
import html
import importlib.util
import io
import logging
import math
import re
from datetime import datetime
from typing import Any, Dict, List, Optional

import httpx
import pandas as pd
import streamlit as st

try:
    import plotly.graph_objects as go
except Exception:  # plotly not installed -> charts fall back to the Genie PNG
    go = None

from genie_client_streamlit import GenieClient, GenieError
from streamlit_runtime import settings
from streamlit_store import store

st.set_page_config(
    page_title="TNS Retail Intelligence",
    page_icon="https://admin.thenewshop.in/static/media/New%20Logo%20.ad69756dd0621a9db47a.jpg",
    layout="wide",
    initial_sidebar_state="expanded",
)

LOGO_URL = "https://admin.thenewshop.in/static/media/New%20Logo%20.ad69756dd0621a9db47a.jpg"
CURRENCY_SYMBOL = "₹"
CHART_COLORS = ["#2563eb", "#f59e0b", "#10b981", "#ef4444", "#8b5cf6"]

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("tns_streamlit")

EXCEL_ENGINE = (
    "xlsxwriter" if importlib.util.find_spec("xlsxwriter")
    else "openpyxl" if importlib.util.find_spec("openpyxl")
    else None
)


# =====================================================================
# CSS
# =====================================================================

def inject_css():
    """Visual-only styling. No application, Genie, data, or rendering logic is changed here."""
    st.markdown(
        """
        <style>
        /* ================================================================
           TNS RETAIL INTELLIGENCE — UI ONLY
           Keep all application behavior untouched. This section controls
           spacing, alignment, typography, colors and Streamlit chrome.
           ================================================================ */
        #MainMenu, footer {visibility:hidden;}
        header {background:transparent !important;}
        .stApp {
            background:#f5f7fb !important;
            color:#172033 !important;
            color-scheme:light !important;
        }

        /* ---------- main page geometry ---------- */
        [data-testid="stMainBlockContainer"],
        .block-container {
            width:100% !important;
            max-width:1120px !important;
            padding-top:20px !important;
            padding-left:28px !important;
            padding-right:28px !important;
            padding-bottom:110px !important;
        }
        [data-testid="stMain"] {background:#f5f7fb !important;}
        section.main {background:#f5f7fb !important;}

        /* ---------- sidebar ---------- */
        [data-testid="stSidebar"] {
            background:#101827 !important;
            min-width:300px !important;
            max-width:300px !important;
            border-right:1px solid #1f2937 !important;
        }
        [data-testid="stSidebar"] > div:first-child {padding:22px 18px 18px !important;}
        [data-testid="stSidebar"] * {color:#dbe3ee;}
        .tns-brand {
            display:flex !important;
            align-items:center !important;
            gap:12px !important;
            padding:4px 0 22px !important;
        }
        .tns-brand img {
            width:46px !important;
            height:46px !important;
            flex:0 0 46px !important;
            border-radius:12px !important;
            object-fit:contain !important;
            background:#fff !important;
            border:1px solid rgba(255,255,255,.12) !important;
        }
        .tns-brand-name {
            font-size:17px !important;
            font-weight:700 !important;
            line-height:1.15 !important;
            color:#fff !important;
            white-space:nowrap !important;
        }
        .section-label {
            font-size:11px !important;
            font-weight:700 !important;
            text-transform:uppercase !important;
            letter-spacing:.8px !important;
            color:#91a0b5 !important;
            margin:18px 0 9px !important;
        }
        [data-testid="stSidebar"] .stButton {margin:0 !important;}
        [data-testid="stSidebar"] .stButton button {
            width:100% !important;
            min-height:42px !important;
            padding:8px 13px !important;
            display:flex !important;
            align-items:center !important;
            justify-content:flex-start !important;
            text-align:left !important;
            background:#1d2939 !important;
            color:#f8fafc !important;
            border:1px solid #334155 !important;
            border-radius:10px !important;
            box-shadow:none !important;
            overflow:hidden !important;
        }
        [data-testid="stSidebar"] .stButton button:hover {
            background:#273449 !important;
            border-color:#4b5d74 !important;
        }
        [data-testid="stSidebar"] .stButton button p {
            width:100% !important;
            overflow:hidden !important;
            text-overflow:ellipsis !important;
            white-space:nowrap !important;
            color:#f8fafc !important;
            font-size:13px !important;
            margin:0 !important;
        }
        [data-testid="stSidebar"] .stButton button[kind="primary"],
        [data-testid="stSidebar"] .stButton button[data-testid="stBaseButton-primary"] {
            background:#39475a !important;
            border-color:#637187 !important;
        }
        [data-testid="stSidebar"] [data-testid="column"] {padding:0 !important;}
        [data-testid="stSidebar"] [data-testid="column"]:last-child .stButton button {
            width:40px !important;
            min-width:40px !important;
            padding:0 !important;
            justify-content:center !important;
            text-align:center !important;
        }
        [data-testid="stSidebar"] [data-testid="column"]:last-child .stButton button p {
            text-align:center !important;
            font-size:13px !important;
        }
        [data-testid="stSidebar"] hr {
            border-color:#334155 !important;
            margin:26px 0 18px !important;
        }
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {
            color:#9eacbf !important;
            font-size:13px !important;
        }

        /* ---------- top conversation header ---------- */
        .tns-header {
            width:100% !important;
            min-height:72px !important;
            display:flex !important;
            flex-direction:column !important;
            justify-content:center !important;
            padding:0 22px !important;
            margin:0 0 26px !important;
            background:#fff !important;
            border:1px solid #e5e9ef !important;
            border-radius:0 !important;
            box-shadow:0 1px 2px rgba(15,23,42,.025) !important;
        }
        .tns-header-title {
            color:#172033 !important;
            font-size:17px !important;
            font-weight:700 !important;
            line-height:1.25 !important;
        }
        .tns-header-subtitle {
            color:#98a2b3 !important;
            font-size:12px !important;
            line-height:1.3 !important;
            margin-top:5px !important;
        }

        /* ---------- welcome ---------- */
        .welcome {
            max-width:720px !important;
            text-align:center !important;
            margin:150px auto 120px !important;
        }
        .welcome h1 {
            color:#172033 !important;
            font-size:31px !important;
            font-weight:750 !important;
            letter-spacing:-.7px !important;
            margin:0 0 12px !important;
        }
        .welcome p {
            color:#7b8798 !important;
            font-size:14px !important;
            margin:0 !important;
        }

        /* ---------- chat messages ---------- */
        [data-testid="stChatMessage"] {
            width:100% !important;
            max-width:100% !important;
            background:transparent !important;
            border:0 !important;
            padding:0 !important;
            margin:0 0 20px !important;
            gap:0 !important;
        }
        [data-testid^="stChatMessageAvatar"] {display:none !important;}

        /* User: compact right-aligned bubble, like the FastAPI UI. */
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {
            width:fit-content !important;
            max-width:78% !important;
            margin-left:auto !important;
            margin-right:0 !important;
            padding:0 !important;
            background:transparent !important;
        }
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) > div:last-child {
            width:fit-content !important;
            max-width:100% !important;
            padding:11px 17px !important;
            background:#111827 !important;
            border:1px solid #111827 !important;
            border-radius:15px 15px 4px 15px !important;
            box-shadow:0 2px 5px rgba(15,23,42,.08) !important;
        }
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) p,
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) span,
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) div {
            color:#fff !important;
            font-size:14px !important;
            line-height:1.5 !important;
        }

        /* Assistant: broad white response card. */
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) {
            width:100% !important;
            background:#fff !important;
            border:1px solid #e1e6ed !important;
            border-radius:15px !important;
            padding:18px 20px !important;
            box-shadow:0 2px 7px rgba(15,23,42,.035) !important;
            overflow:visible !important;
        }
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) p,
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) li {
            color:#263244 !important;
            font-size:14px !important;
            line-height:1.65 !important;
        }
        [data-testid="stChatMessage"] h1 {font-size:21px !important;margin:12px 0 8px !important;color:#172033 !important;}
        [data-testid="stChatMessage"] h2 {font-size:18px !important;margin:14px 0 7px !important;color:#172033 !important;}
        [data-testid="stChatMessage"] h3 {font-size:16px !important;margin:13px 0 6px !important;color:#172033 !important;}
        [data-testid="stChatMessage"] h4 {font-size:14px !important;margin:11px 0 5px !important;color:#172033 !important;}
        [data-testid="stChatMessage"] table {
            width:100% !important;
            border-collapse:separate !important;
            border-spacing:0 !important;
            overflow:hidden !important;
            border:1px solid #e5e7eb !important;
            border-radius:9px !important;
            font-size:13px !important;
            margin:10px 0 12px !important;
        }
        [data-testid="stChatMessage"] th,
        [data-testid="stChatMessage"] td {
            border-right:1px solid #e5e7eb !important;
            border-bottom:1px solid #e5e7eb !important;
            padding:8px 10px !important;
            color:#263244 !important;
            background:#fff !important;
        }
        [data-testid="stChatMessage"] th {
            background:#f7f9fc !important;
            font-weight:700 !important;
        }
        [data-testid="stChatMessage"] sup {font-size:10px !important;color:#2563eb !important;font-weight:700 !important;}

        /* ---------- chart / table cards ---------- */
        [data-testid="stChatMessage"] [data-testid="stVerticalBlockBorderWrapper"] {
            width:100% !important;
            background:#fff !important;
            border:1px solid #dfe5ec !important;
            border-radius:12px !important;
            padding:0 !important;
            overflow:hidden !important;
            box-shadow:none !important;
            margin:10px 0 16px !important;
        }
        [data-testid="stChatMessage"] [data-testid="stVerticalBlockBorderWrapper"] > div {
            padding:0 !important;
        }
        [data-testid="stPlotlyChart"] {
            width:100% !important;
            margin:0 !important;
            padding:0 !important;
        }
        [data-testid="stPlotlyChart"] iframe {border-radius:0 !important;}

        /* ---------- main-area buttons / downloads ---------- */
        [data-testid="stMain"] .stButton button,
        [data-testid="stMain"] .stDownloadButton button,
        section.main .stButton button,
        section.main .stDownloadButton button {
            min-height:34px !important;
            padding:6px 13px !important;
            border:1px solid #cfd6df !important;
            border-radius:8px !important;
            background:#fff !important;
            color:#344054 !important;
            font-size:12px !important;
            font-weight:550 !important;
            box-shadow:none !important;
        }
        [data-testid="stMain"] .stButton button:hover,
        [data-testid="stMain"] .stDownloadButton button:hover {
            background:#f7f9fc !important;
            border-color:#aeb8c6 !important;
            color:#172033 !important;
        }
        [data-testid="stMain"] .stButton button p,
        [data-testid="stMain"] .stDownloadButton button p {
            font-size:12px !important;
            margin:0 !important;
        }
        /* Align button rows instead of allowing Streamlit's default vertical padding to drift. */
        [data-testid="stChatMessage"] [data-testid="column"] {
            display:flex !important;
            align-items:flex-start !important;
        }
        [data-testid="stChatMessage"] [data-testid="column"] .stButton,
        [data-testid="stChatMessage"] [data-testid="column"] .stDownloadButton {
            width:100% !important;
            margin:0 !important;
        }

        /* ---------- thought process ---------- */
        [data-testid="stExpander"] {
            border:1px solid #e2e7ee !important;
            border-radius:10px !important;
            background:#fff !important;
            margin:4px 0 14px !important;
            box-shadow:none !important;
        }
        [data-testid="stExpander"] summary {
            background:#fff !important;
            color:#344054 !important;
            min-height:42px !important;
            padding:0 13px !important;
        }
        [data-testid="stExpander"] summary:hover {background:#f8fafc !important;}
        [data-testid="stExpander"] summary p {font-size:13px !important;color:#344054 !important;}

        /* ---------- login ---------- */
        [data-testid="stForm"] {
            width:100% !important;
            max-width:480px !important;
            margin:11vh auto 0 !important;
            padding:34px 36px 30px !important;
            background:#fff !important;
            border:1px solid #e2e7ee !important;
            border-radius:18px !important;
            box-shadow:0 18px 50px rgba(15,23,42,.08) !important;
        }
        [data-testid="stForm"] [data-testid="stTextInput"] label,
        [data-testid="stForm"] [data-testid="stTextInput"] label p {
            color:#344054 !important;
            font-size:13px !important;
            font-weight:600 !important;
        }
        [data-testid="stForm"] [data-testid="stTextInput"] input {
            height:44px !important;
            color:#111827 !important;
            -webkit-text-fill-color:#111827 !important;
            background:#fff !important;
            border:1px solid #cfd6df !important;
            border-radius:9px !important;
        }
        [data-testid="stForm"] button {
            min-height:44px !important;
            color:#fff !important;
            -webkit-text-fill-color:#fff !important;
            background:#111827 !important;
            border:1px solid #111827 !important;
            border-radius:9px !important;
            font-weight:650 !important;
        }
        .login-logo {
            display:block !important;
            width:64px !important;
            height:64px !important;
            object-fit:contain !important;
            margin:0 auto 16px !important;
            border-radius:12px !important;
            background:#fff !important;
            border:1px solid #e5e7eb !important;
        }
        .login-title {text-align:center !important;font-size:25px !important;font-weight:750 !important;color:#172033 !important;margin:0 0 7px !important;}
        .login-sub {text-align:center !important;color:#667085 !important;font-size:14px !important;margin:0 0 28px !important;}
        .login-heading {text-align:center !important;}

        /* ---------- fixed chat composer ---------- */
        [data-testid="stBottom"], [data-testid="stBottom"] > div {
            background:rgba(245,247,251,.96) !important;
            backdrop-filter:blur(10px) !important;
        }
        [data-testid="stBottomBlockContainer"] {
            max-width:1120px !important;
            padding:10px 28px 16px !important;
        }
        [data-testid="stChatInput"] {
            width:100% !important;
            background:#fff !important;
            border:1px solid #cfd6df !important;
            border-radius:14px !important;
            box-shadow:0 5px 18px rgba(15,23,42,.08) !important;
            min-height:62px !important;
        }
        [data-testid="stChatInput"] > div {
            background:#fff !important;
            border:0 !important;
        }
        [data-testid="stChatInput"] textarea,
        [data-testid="stChatInput"] input,
        [data-testid="stChatInput"] textarea:focus,
        [data-testid="stChatInput"] input:focus {
            min-height:46px !important;
            color:#111827 !important;
            -webkit-text-fill-color:#111827 !important;
            caret-color:#111827 !important;
            background:#fff !important;
            font-size:15px !important;
        }
        [data-testid="stChatInput"] textarea::placeholder,
        [data-testid="stChatInput"] input::placeholder {
            color:#7b8798 !important;
            -webkit-text-fill-color:#7b8798 !important;
            opacity:1 !important;
        }
        [data-testid="stChatInput"] button {
            width:42px !important;
            height:42px !important;
            min-width:42px !important;
            margin-right:5px !important;
            border-radius:10px !important;
            background:#111827 !important;
            border:1px solid #111827 !important;
        }
        [data-testid="stChatInput"] button:hover {background:#1f2937 !important;}
        [data-testid="stChatInput"] button svg {fill:#fff !important;color:#fff !important;}

        /* ---------- responsive ---------- */
        @media (max-width: 900px) {
            [data-testid="stSidebar"] {min-width:260px !important;max-width:260px !important;}
            [data-testid="stMainBlockContainer"], .block-container {
                max-width:100% !important;
                padding-left:16px !important;
                padding-right:16px !important;
            }
            [data-testid="stBottomBlockContainer"] {padding-left:16px !important;padding-right:16px !important;}
            [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {max-width:88% !important;}
        }
        @media (max-width: 640px) {
            [data-testid="stSidebar"] {min-width:240px !important;max-width:240px !important;}
            .tns-brand-name {font-size:15px !important;}
            .tns-header {padding:0 16px !important;}
            .welcome {margin:100px auto 80px !important;}
            .welcome h1 {font-size:26px !important;}
            [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {max-width:94% !important;}
            [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) {padding:14px !important;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =====================================================================
# Streamlit version helpers
# =====================================================================

_WIDE_MODE: Dict[str, str] = {}


def wide(fn, *args, **kwargs):
    """Call a Streamlit element stretched to full width, on old and new versions."""
    name = getattr(fn, "__name__", str(fn))
    mode = _WIDE_MODE.get(name)
    if mode is None:
        try:
            result = fn(*args, width="stretch", **kwargs)
            _WIDE_MODE[name] = "stretch"
            return result
        except Exception:
            _WIDE_MODE[name] = "legacy"
            return fn(*args, use_container_width=True, **kwargs)
    if mode == "stretch":
        return fn(*args, width="stretch", **kwargs)
    return fn(*args, use_container_width=True, **kwargs)


# =====================================================================
# Genie call plumbing (unchanged behaviour)
# =====================================================================

def run_async(coro):
    return asyncio.run(coro)


def new_client() -> GenieClient:
    return GenieClient()


async def _call_genie(coro_factory):
    client = new_client()
    try:
        return await coro_factory(client)
    finally:
        await client.close()


def call_genie(coro_factory):
    return run_async(_call_genie(coro_factory))


def check_credentials(username: str, password: str) -> bool:
    return username == settings.app_username and password == settings.app_password


def ensure_store():
    if st.session_state.get("store_initialized"):
        return
    store.initialize()
    st.session_state.store_initialized = True


# =====================================================================
# Number / cell formatting (Indian style, same rules as the web UI)
# =====================================================================

_NUMERIC_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")


def normalize_name(name: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name or "").lower())


def is_numeric_value(value: Any) -> bool:
    if value is None or value == "":
        return False
    text = str(value).strip()
    if not _NUMERIC_RE.match(text):
        return False
    try:
        return math.isfinite(float(text))
    except ValueError:
        return False


def is_quantity_column(name: Any) -> bool:
    n = normalize_name(name)
    words = ("quantity", "qty", "units", "unitcount", "itemcount", "productcount",
             "skucount", "ordercount", "storecount", "customercount")
    return any(w in n for w in words) or n.endswith("count")


def is_percentage_column(name: Any) -> bool:
    n = normalize_name(name)
    return any(w in n for w in ("percent", "percentage", "rate", "margin")) or n.endswith("pct")


def is_currency_column(name: Any) -> bool:
    if is_quantity_column(name) or is_percentage_column(name):
        return False
    n = normalize_name(name)
    words = ("revenue", "sales", "mrp", "amount", "price", "cost", "profit", "gmv",
             "turnover", "discount", "deliveryfee", "taxamount", "netamount", "grossamount")
    return any(w in n for w in words)


def indian_format(value: float, min_dec: int = 0, max_dec: int = 2) -> str:
    """Format a non-negative number with Indian digit grouping (12,34,567.89)."""
    text = f"{abs(value):.{max_dec}f}"
    int_part, _, dec = text.partition(".")
    dec = dec.rstrip("0")
    if len(dec) < min_dec:
        dec = dec.ljust(min_dec, "0")
    if len(int_part) > 3:
        head, tail = int_part[:-3], int_part[-3:]
        groups: List[str] = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        int_part = ",".join(groups + [tail])
    return int_part + (f".{dec}" if dec else "")


def format_indian_currency(value: Any) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if not math.isfinite(number):
        return str(value)
    sign = "-" if number < 0 else ""
    absolute = abs(number)
    if absolute >= 10_000_000:
        return f"{sign}{CURRENCY_SYMBOL}{indian_format(absolute / 10_000_000, 2, 2)} Cr"
    if absolute >= 100_000:
        return f"{sign}{CURRENCY_SYMBOL}{indian_format(absolute / 100_000, 2, 2)} L"
    return f"{sign}{CURRENCY_SYMBOL}{indian_format(absolute, 0, 2)}"


def format_plain_number(value: Any) -> str:
    number = float(value)
    return ("-" if number < 0 else "") + indian_format(number, 0, 2)


def format_cell(value: Any, column: str = "") -> str:
    if value is None or value == "":
        return "—"
    text = str(value).strip()
    if not is_numeric_value(text):
        return str(value)
    if is_quantity_column(column):
        return format_plain_number(text)
    if is_percentage_column(column):
        return f"{format_plain_number(text)}%"
    if is_currency_column(column):
        return format_indian_currency(text)
    return format_plain_number(text)


def column_kind(name: Any) -> str:
    if is_currency_column(name):
        return "currency"
    if is_quantity_column(name):
        return "quantity"
    if is_percentage_column(name):
        return "percent"
    return "number"


def format_chart_value(value: Any, kind: str) -> str:
    if value is None:
        return "—"
    if kind == "currency":
        return format_indian_currency(value)
    formatted = format_plain_number(value)
    return f"{formatted}%" if kind == "percent" else formatted


def unique_columns(columns: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    result = []
    for column in columns:
        count = seen.get(column, 0)
        seen[column] = count + 1
        result.append(column if count == 0 else f"{column} ({count + 1})")
    return result


def normalize_rows(columns: List[str], rows: List[Any]) -> List[List[Any]]:
    width = len(columns)
    out = []
    for row in rows:
        row = list(row) if isinstance(row, (list, tuple)) else [row]
        out.append((row + [None] * width)[:width])
    return out


# =====================================================================
# Downloads (CSV / Excel / full CSV)
# =====================================================================

def file_safe_name(value: Any) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", " ", str(value or "table"))
    name = re.sub(r"\s+", "-", name).strip("-")[:80]
    return name or "table"


def csv_bytes(columns: List[str], rows: List[List[Any]]) -> bytes:
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\r\n")
    writer.writerow(columns)
    for row in rows:
        writer.writerow(["" if v is None else v for v in row])
    # BOM so Excel opens UTF-8 (₹ etc.) correctly.
    return ("\ufeff" + buffer.getvalue()).encode("utf-8")


def coerce_excel_cell(value: Any, allow_formatted: bool = False) -> Any:
    if value is None or value == "":
        return ""
    candidate = str(value).strip()
    if allow_formatted:
        candidate = re.sub(r"^₹\s?", "", candidate).replace(",", "")
    if (
        is_numeric_value(candidate)
        and not re.match(r"^[+-]?0\d", candidate)
        and len(re.sub(r"\D", "", candidate)) <= 15
    ):
        return float(candidate)
    return value


@st.cache_data(show_spinner=False, max_entries=64)
def excel_bytes(columns: tuple, rows: tuple, sheet_name: str, allow_formatted: bool) -> Optional[bytes]:
    if EXCEL_ENGINE is None:
        return None
    cols = unique_columns(list(columns))
    data = [[coerce_excel_cell(c, allow_formatted) for c in row] for row in rows]
    frame = pd.DataFrame(data, columns=cols)
    buffer = io.BytesIO()
    safe_sheet = re.sub(r"[\\/?*\[\]:]", "", sheet_name)[:31] or "Sheet1"
    with pd.ExcelWriter(buffer, engine=EXCEL_ENGINE) as writer:
        frame.to_excel(writer, index=False, sheet_name=safe_sheet)
    return buffer.getvalue()


def csv_button(columns, rows, base_name, key):
    wide(
        st.download_button,
        "⬇ CSV",
        data=csv_bytes(columns, rows),
        file_name=f"{base_name}.csv",
        mime="text/csv",
        key=f"csv_{key}",
    )


def excel_button(columns, rows, base_name, key, allow_formatted=False):
    if EXCEL_ENGINE is None:
        return
    data = excel_bytes(tuple(columns), tuple(tuple(r) for r in rows), base_name, allow_formatted)
    if not data:
        return
    wide(
        st.download_button,
        "⬇ Excel",
        data=data,
        file_name=f"{base_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"xlsx_{key}",
    )


def fetch_full_csv(conversation_id: str, message_id: str, attachment_id: str) -> bytes:
    links = call_genie(
        lambda client: client.get_full_query_download_links(conversation_id, message_id, attachment_id)
    )
    parts: List[bytes] = []
    for link in links:
        response = httpx.get(link, timeout=300.0, follow_redirects=True)
        response.raise_for_status()
        parts.append(response.content)
    if not parts:
        raise GenieError("Databricks returned no download files.")

    merged = parts[0].lstrip(b"\xef\xbb\xbf")
    header = merged.split(b"\n", 1)[0].strip()
    for extra in parts[1:]:
        extra = extra.lstrip(b"\xef\xbb\xbf")
        first = extra.split(b"\n", 1)[0].strip()
        if first == header and b"\n" in extra:
            extra = extra.split(b"\n", 1)[1]
        if not merged.endswith(b"\n"):
            merged += b"\n"
        merged += extra
    return b"\xef\xbb\xbf" + merged


def full_csv_control(table: Dict[str, Any], conversation_id: str, message_id: str, key: str):
    attachment_id = table.get("attachment_id")
    if not attachment_id or not message_id or not conversation_id:
        return
    state_key = f"fullcsv_{key}"
    data = st.session_state.get(state_key)
    if data is not None:
        wide(
            st.download_button,
            "⬇ Full CSV",
            data=data,
            file_name=f"{file_safe_name(table.get('title') or 'TNS-result')}-full.csv",
            mime="text/csv",
            key=f"fullcsv_dl_{key}",
        )
        return
    if wide(st.button, "Full CSV (all rows)", key=f"fullcsv_btn_{key}"):
        try:
            with st.spinner("Preparing full CSV…"):
                st.session_state[state_key] = fetch_full_csv(conversation_id, message_id, attachment_id)
            st.rerun()
        except GenieError as exc:
            st.warning(f"Unable to prepare the full CSV: {exc}")
        except httpx.HTTPError as exc:
            st.warning(f"Unable to download the full CSV: {exc}")


# =====================================================================
# Markdown preparation (citations, ₹, glued headings, tables)
# =====================================================================

_CITATION_RE = re.compile(
    r"\[unrendered\s+:citation\[([^\]]+)\]\]|:citation\[([^\]]+)\](?:\{[^}]*\})?"
)
_OTHER_DIRECTIVE_RE = re.compile(r"\[unrendered\s+:[A-Za-z]+\[[^\]]*\]\]")
_GLUED_HEADING_MARKER_RE = re.compile(r"([A-Za-z0-9)\.:\]])(#{2,6})[ \t]+(?=[A-Z])")
_GLUED_HEADING_LINE_RE = re.compile(r"^(?:#{1,6}\s+)?((?:[A-Z][a-z]+ ){0,3}[A-Z][a-z]+)(?=[A-Z][a-z]+ [a-z])")
_TABLE_SEPARATOR_RE = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$")


def normalize_source_id(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def prepare_text(text: str, source_titles: Optional[Dict[str, str]] = None):
    """Return (markdown, cited_ids) ready for st.markdown(unsafe_allow_html=True)."""
    source_titles = source_titles or {}
    cited: List[str] = []

    def citation(match: "re.Match[str]") -> str:
        raw_id = match.group(1) or match.group(2) or ""
        cid = re.sub(r"[^A-Za-z0-9_-]", "", raw_id.strip())
        if not cid:
            return ""
        if cid not in cited:
            cited.append(cid)
        number = cited.index(cid) + 1
        title = html.escape(source_titles.get(normalize_source_id(cid), ""), quote=True)
        tip = f' title="{title}"' if title else ""
        return f"<sup{tip}>[{number}]</sup>"

    source = str(text or "")
    source = source.replace("<", "&lt;")
    source = _CITATION_RE.sub(citation, source)
    source = _OTHER_DIRECTIVE_RE.sub("", source)
    source = re.sub(r"\$(?=\s?\d)", CURRENCY_SYMBOL, source)
    # Any remaining "$" would start LaTeX in Streamlit's markdown.
    source = source.replace("$", "\\$")
    source = _GLUED_HEADING_MARKER_RE.sub(r"\1\n\n\2 ", source)

    lines: List[str] = []
    for raw in source.split("\n"):
        trimmed = raw.strip()
        if len(trimmed) > 60 and not re.match(r"^[|\-*\d]", trimmed):
            match = _GLUED_HEADING_LINE_RE.match(trimmed)
            if match and len(match.group(1)) <= 45:
                rest = trimmed[trimmed.index(match.group(1)) + len(match.group(1)):]
                lines.extend([f"### {match.group(1)}", "", rest])
                continue
        lines.append(raw)
    return "\n".join(lines), cited


def split_table_row(row: str) -> List[str]:
    row = row.strip()
    row = row[1:] if row.startswith("|") else row
    row = row[:-1] if row.endswith("|") else row
    return [cell.strip() for cell in row.split("|")]


def split_blocks(text: str) -> List[Dict[str, Any]]:
    lines = text.split("\n")
    blocks: List[Dict[str, Any]] = []
    buffer: List[str] = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("|") and i + 1 < len(lines) and _TABLE_SEPARATOR_RE.match(lines[i + 1].strip()):
            if buffer:
                blocks.append({"type": "text", "text": "\n".join(buffer)})
                buffer = []
            raw = [lines[i], lines[i + 1]]
            head = split_table_row(line)
            body: List[List[str]] = []
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                raw.append(lines[i])
                body.append(split_table_row(lines[i]))
                i += 1
            blocks.append({"type": "table", "raw": raw, "head": head, "body": body})
            continue
        buffer.append(lines[i])
        i += 1
    if buffer:
        blocks.append({"type": "text", "text": "\n".join(buffer)})
    return blocks


def clean_cell(cell: str) -> str:
    cell = re.sub(r"<[^>]+>", "", cell)
    cell = re.sub(r"\[([^\]]+)\]\((?:https?://)[^)\s]+\)", r"\1", cell)
    cell = cell.replace("**", "").replace("__", "").replace("`", "").replace("\\$", "$")
    return html.unescape(cell).strip()


def render_answer(content: str, uid: str, source_titles: Optional[Dict[str, str]] = None) -> List[str]:
    text, cited = prepare_text(content, source_titles)
    for index, block in enumerate(split_blocks(text)):
        if block["type"] == "text":
            if block["text"].strip():
                st.markdown(block["text"], unsafe_allow_html=True)
            continue

        st.markdown("\n".join(block["raw"]), unsafe_allow_html=True)
        columns = [clean_cell(c) for c in block["head"]]
        rows = normalize_rows(columns, [[clean_cell(c) for c in row] for row in block["body"]])
        base = f"TNS-table-{index + 1}"
        slots = st.columns([1, 1, 5] if EXCEL_ENGINE else [1, 6])
        with slots[0]:
            csv_button(columns, rows, base, f"{uid}_md{index}")
        if EXCEL_ENGINE:
            with slots[1]:
                excel_button(columns, rows, base, f"{uid}_md{index}", allow_formatted=True)
    return cited


# =====================================================================
# Tables
# =====================================================================

def render_table(table: Dict[str, Any], conversation_id: str, message_id: str, key: str):
    columns = [str(c) for c in (table.get("columns") or [])]
    rows = normalize_rows(columns, table.get("rows") or []) if columns else []
    title = table.get("title") or "Query result"
    base = file_safe_name(title)
    displayed = table.get("displayed_row_count", len(rows))
    total = int(table.get("row_count") or displayed or 0)
    truncated = bool(table.get("truncated")) or total > displayed

    with st.container(border=True):
        st.markdown(f"**{title}**")
        meta = f"{indian_format(total, 0, 0)} row(s)"
        if truncated:
            meta += f" · showing first {indian_format(displayed, 0, 0)}"
        st.caption(meta)
        if table.get("description"):
            st.caption(table["description"])

        if not columns:
            st.info(table.get("error") or "No tabular result was returned.")
            return

        slot_widths = [1, 1, 1.9, 1.3, 2.4]
        slots = st.columns(slot_widths)
        with slots[0]:
            if rows:
                csv_button(columns, rows, base, key)
        with slots[1]:
            if rows:
                excel_button(columns, rows, base, key)
        with slots[2]:
            if truncated:
                full_csv_control(table, conversation_id, message_id, key)
        show_code = False
        with slots[3]:
            if table.get("query"):
                show_code = st.toggle("Show code", key=f"code_{key}")

        if show_code:
            st.code(table["query"], language="sql")

        frame = pd.DataFrame(
            [[format_cell(v, c) for v, c in zip(row, columns)] for row in rows],
            columns=unique_columns(columns),
        )
        height = int(min(440, 38 * (len(rows) + 1) + 3))
        wide(st.dataframe, frame, hide_index=True, height=height)

        if truncated:
            st.caption("Only the first rows are shown here; use Full CSV for the complete TNS query result.")
        if table.get("error"):
            st.caption(str(table["error"]))


# =====================================================================
# Charts
# =====================================================================

_SERIES_GROUPS = [
    ["quantity", "qty", "units", "unit"],
    ["revenue", "sales", "amount", "value"],
    ["orders", "order", "transactions", "transaction", "invoices", "invoice", "bills"],
]


def _series_group(name: Any) -> int:
    text = normalize_name(name)
    for index, words in enumerate(_SERIES_GROUPS):
        if any(w in text for w in words):
            return index
    return -1


def is_iso_date(value: Any) -> bool:
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}([T ]|$)", str(value)))


def format_chart_label(value: Any, monthly: bool) -> str:
    if not is_iso_date(value):
        return str(value)
    try:
        date = datetime.strptime(str(value)[:10], "%Y-%m-%d")
    except ValueError:
        return str(value)
    return date.strftime("%b %Y") if monthly else f"{date.day} {date.strftime('%b %Y')}"


def find_source_table(viz: Dict[str, Any], tables: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    for attachment_id in (viz.get("query_attachment_id"), viz.get("attachment_id")):
        if not attachment_id:
            continue
        for table in tables:
            if table.get("attachment_id") == attachment_id:
                return table
    return tables[-1] if tables else None


def build_chart_spec(table: Dict[str, Any], viz: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    columns = [str(c) for c in (table.get("columns") or [])]
    rows = normalize_rows(columns, (table.get("rows") or [])[:50]) if columns else []
    if len(columns) < 2 or not rows:
        return None

    numeric_columns = []
    for c in range(1, len(columns)):
        values = [row[c] for row in rows if row[c] not in (None, "")]
        if values and all(is_numeric_value(v) for v in values):
            numeric_columns.append(c)
    if not numeric_columns:
        return None

    date_axis = all(is_iso_date(row[0]) for row in rows)
    monthly = date_axis and all(re.match(r"^\d{4}-\d{2}-01", str(row[0])) for row in rows)
    hint = " ".join(str(viz.get(k) or "") for k in ("chart_type", "type", "title")).lower()
    chart_type = "line" if (date_axis or re.search(r"line|trend", hint)) else "bar"

    title_text = str(viz.get("title") or "").lower()
    title_groups = [
        group
        for group, words in enumerate(_SERIES_GROUPS)
        if any(re.search(rf"\b{w}s?\b", title_text) for w in words)
    ]
    preferred = [c for c in numeric_columns if _series_group(columns[c]) in title_groups]
    if not preferred:
        first_kind = column_kind(columns[numeric_columns[0]])
        preferred = [c for c in numeric_columns if column_kind(columns[c]) == first_kind]

    preferred_set = set(preferred)
    first_kind = column_kind(columns[preferred[0]])
    ordered = preferred + [c for c in numeric_columns if c not in preferred_set]

    datasets = [
        {
            "label": columns[c],
            "kind": column_kind(columns[c]),
            "hidden": c not in preferred_set,
            "data": [float(row[c]) if is_numeric_value(row[c]) else None for row in rows],
        }
        for c in ordered[:5]
    ]
    return {
        "type": chart_type,
        "labels": [format_chart_label(row[0], monthly) for row in rows],
        "datasets": datasets,
        "first_kind": first_kind,
        "horizontal": chart_type == "bar" and len(rows) > 12,
    }


def nice_ticks(vmin: float, vmax: float, target: int = 5) -> Optional[List[float]]:
    if not (math.isfinite(vmin) and math.isfinite(vmax)) or vmax <= vmin:
        return None
    raw = (vmax - vmin) / target
    magnitude = 10 ** math.floor(math.log10(raw))
    step = magnitude * 10
    for multiplier in (1, 2, 2.5, 5, 10):
        if raw <= multiplier * magnitude:
            step = multiplier * magnitude
            break
    start = math.floor(vmin / step) * step
    end = math.ceil(vmax / step) * step
    count = min(int(round((end - start) / step)) + 1, 14)
    return [round(start + i * step, 10) for i in range(count)]


def build_figure(spec: Dict[str, Any]):
    horizontal = spec["horizontal"]
    labels = spec["labels"]
    fig = go.Figure()

    for index, dataset in enumerate(spec["datasets"]):
        color = CHART_COLORS[index % len(CHART_COLORS)]
        hover = [
            f"{label}<br>{dataset['label']}: {format_chart_value(value, dataset['kind'])}"
            for label, value in zip(labels, dataset["data"])
        ]
        common = dict(
            name=dataset["label"],
            visible="legendonly" if dataset["hidden"] else True,
            hovertext=hover,
            hovertemplate="%{hovertext}<extra></extra>",
        )
        if spec["type"] == "line":
            fig.add_trace(go.Scatter(
                x=labels, y=dataset["data"], mode="lines+markers",
                line=dict(color=color, width=2), marker=dict(color=color, size=6),
                connectgaps=False, **common,
            ))
        elif horizontal:
            fig.add_trace(go.Bar(y=labels, x=dataset["data"], orientation="h",
                                 marker_color=color, **common))
        else:
            fig.add_trace(go.Bar(x=labels, y=dataset["data"], marker_color=color, **common))

    visible_values = [
        v for ds in spec["datasets"] if not ds["hidden"] for v in ds["data"] if v is not None
    ]
    value_axis: Dict[str, Any] = dict(gridcolor="#eef0f3", zeroline=True, zerolinecolor="#d1d5db")
    if visible_values:
        ticks = nice_ticks(min(0.0, min(visible_values)), max(0.0, max(visible_values)))
        if ticks:
            value_axis.update(
                tickvals=ticks,
                ticktext=[format_chart_value(t, spec["first_kind"]) for t in ticks],
            )
    category_axis: Dict[str, Any] = dict(
        type="category", categoryorder="array", categoryarray=labels,
        automargin=True, showgrid=False,
    )

    height = max(340, len(labels) * 24 + 90) if horizontal else 340
    visible_count = sum(1 for dataset in spec["datasets"] if not dataset["hidden"])
    fig.update_layout(
        template="plotly_white",
        height=max(380, height),
        margin=dict(l=62, r=24, t=58, b=62),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        showlegend=visible_count > 1 or len(spec["datasets"]) > 1,
        legend=dict(
            orientation="h",
            yanchor="bottom", y=1.02,
            xanchor="left", x=0,
            bgcolor="rgba(255,255,255,0)",
            borderwidth=0,
            font=dict(size=12, color="#667085"),
            itemclick="toggle",
            itemdoubleclick="toggleothers",
        ),
        barmode="group",
        bargap=0.28,
        bargroupgap=0.08,
        hovermode="x unified" if spec["type"] == "line" and not horizontal else "closest",
        font=dict(family="Inter, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif", size=12, color="#344054"),
        hoverlabel=dict(bgcolor="#101828", font_color="#ffffff", bordercolor="#101828"),
        dragmode="pan",
    )
    if horizontal:
        fig.update_xaxes(**value_axis)
        fig.update_yaxes(autorange="reversed", **category_axis)
    else:
        fig.update_yaxes(
            **value_axis,
            title_font=dict(size=12, color="#344054"),
            tickfont=dict(size=11, color="#667085"),
        )
        fig.update_xaxes(
            tickangle=-35 if len(labels) > 10 else 0,
            **category_axis,
            tickfont=dict(size=11, color="#667085"),
            showline=True,
            linecolor="#d0d5dd",
        )
    return fig


@st.cache_data(show_spinner=False, ttl=3600, max_entries=32)
def fetch_visualization_png(conversation_id: str, message_id: str, attachment_id: str) -> bytes:
    return call_genie(
        lambda client: client.download_visualization(conversation_id, message_id, attachment_id)
    )


def render_visualization(
    viz: Dict[str, Any],
    tables: List[Dict[str, Any]],
    conversation_id: str,
    message_id: str,
    key: str,
):
    with st.container(border=True):
        st.markdown(f"**{viz.get('title') or 'Visualization'}**")
        table = find_source_table(viz, tables)

        if go is not None and table:
            try:
                spec = build_chart_spec(table, viz)
                if spec:
                    figure = build_figure(spec)
                    wide(st.plotly_chart, figure, key=f"chart_{key}", config={"displaylogo": False})
                    st.caption(f"Source: {table.get('title') or 'Query result'}")
                    with st.expander("View data"):
                        render_table(table, conversation_id, message_id, f"{key}_data")
                    return
            except Exception as exc:
                logger.warning("Native chart failed, using TNS image: %s", exc)

        attachment_id = viz.get("attachment_id")
        if not attachment_id or not message_id:
            st.info("This visualization is not available.")
            return
        try:
            image = fetch_visualization_png(conversation_id, message_id, attachment_id)
            wide(st.image, image)
            st.download_button(
                "⬇ Chart (PNG)",
                data=image,
                file_name="TNS_visualization.png",
                mime="image/png",
                key=f"viz_dl_{key}",
            )
        except Exception as exc:
            logger.warning("Visualization retrieval failed: %s", exc)
            st.warning("The analytical response was returned, but this visualization could not be loaded.")


# =====================================================================
# Messages
# =====================================================================

def render_thoughts(thoughts: List[Any], uid: str):
    if not thoughts:
        return
    with st.expander("Thought process", expanded=False):
        number = 0
        for thought in thoughts:
            content = thought.get("content") if isinstance(thought, dict) else str(thought)
            if not content:
                continue
            number += 1
            text, _ = prepare_text(content)
            st.markdown(f"**{number}.** {text}", unsafe_allow_html=True)


def presentation_parts(presentation: Dict[str, Any]):
    blocks = presentation.get("blocks") or []

    def from_blocks(block_type: str) -> List[Any]:
        return [b.get("data") for b in blocks if b.get("type") == block_type and b.get("data") is not None]

    thoughts = presentation.get("thoughts") or (from_blocks("thoughts") or [[]])[0]
    tables = presentation.get("tables") or from_blocks("table")
    visualizations = presentation.get("visualizations") or from_blocks("visualization")
    suggested = presentation.get("suggested_questions") or (from_blocks("suggested_questions") or [[]])[0]
    return thoughts, tables, visualizations, suggested


def source_title_map(tables: List[Dict[str, Any]], visualizations: List[Dict[str, Any]]) -> Dict[str, str]:
    titles: Dict[str, str] = {}
    for table in tables:
        if table.get("attachment_id"):
            titles[normalize_source_id(table["attachment_id"])] = table.get("title") or "Query result"
    for viz in visualizations:
        source = find_source_table(viz, tables)
        label = (source or {}).get("title") or viz.get("title") or "Visualization"
        for field in ("attachment_id", "query_attachment_id"):
            if viz.get(field):
                titles.setdefault(normalize_source_id(viz[field]), label)
    return titles


def render_assistant(content: str, presentation: Dict[str, Any], conversation_id: str,
                     message_id: str, uid: str, show_suggestions: bool):
    thoughts, tables, visualizations, suggested = presentation_parts(presentation or {})
    titles = source_title_map(tables, visualizations)

    # Layout mirrors the web UI: thought process, report, charts, tables, suggestions.
    render_thoughts(thoughts, uid)
    cited = render_answer(content, uid, titles)

    sources = [
        f"[{number}] {titles[normalize_source_id(cid)]}"
        for number, cid in enumerate(cited, start=1)
        if normalize_source_id(cid) in titles
    ]
    if sources:
        st.caption("Sources: " + " · ".join(sources))

    chart_sources = set()
    for index, viz in enumerate(visualizations):
        render_visualization(viz, tables, conversation_id, message_id, f"{uid}_v{index}")
        source = find_source_table(viz, tables)
        if source:
            chart_sources.add(source.get("attachment_id"))

    for index, table in enumerate(tables):
        if table.get("attachment_id") in chart_sources:
            continue
        render_table(table, conversation_id, message_id, f"{uid}_t{index}")

    if suggested and show_suggestions:
        st.markdown('<div class="section-label">Suggested questions</div>', unsafe_allow_html=True)
        for index, question in enumerate(suggested):
            if wide(st.button, question, key=f"suggestion_{uid}_{index}"):
                st.session_state.pending_prompt = question
                st.rerun()


def render_message(message: Dict[str, Any], conversation_id: str, uid: str, is_last: bool):
    role = message.get("role")
    content = message.get("content") or ""
    if role == "user":
        with st.chat_message("user"):
            st.markdown(content.replace("$", "\\$"))
        return

    presentation = message.get("presentation") or message.get("metadata") or {}
    message_id = presentation.get("agent_message_id") or message.get("message_id") or ""
    with st.chat_message("assistant"):
        render_assistant(content, presentation, conversation_id, message_id, uid, show_suggestions=is_last)


# =====================================================================
# Turn processing
# =====================================================================

async def run_agent_turn(client: GenieClient, message: str, conversation_id: Optional[str]):
    original = conversation_id
    rebound = False
    try:
        response = await client.create_agent_response(
            message, conversation_id=conversation_id, enable_visualization=True
        )
    except GenieError as exc:
        if conversation_id and client.is_legacy_conversation_error(exc):
            rebound = True
            response = await client.create_agent_response(
                message, conversation_id=None, enable_visualization=True
            )
        else:
            raise

    resolved = response.get("conversation_id") or (None if rebound else conversation_id)
    if not resolved:
        raise GenieError("TNS Agent did not return a conversation ID.")

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
    sid = store.create_session(result["conversation_id"], st.session_state.username, title)
    store.add_message(sid, st.session_state.username, "user", prompt)
    store.add_message(sid, st.session_state.username, "assistant", result["answer"], result["presentation"])
    st.session_state.active_chat_id = sid
    return result


def process_followup(session_id: str, prompt: str):
    conversation_id = store.get_conversation_id(session_id, st.session_state.username)
    if not conversation_id:
        raise GenieError("Chat session not found.")
    result = call_genie(lambda client: run_agent_turn(client, prompt, conversation_id))
    if result["conversation_changed"]:
        store.set_conversation_id(session_id, st.session_state.username, result["conversation_id"])
    store.add_message(session_id, st.session_state.username, "user", prompt)
    store.add_message(session_id, st.session_state.username, "assistant", result["answer"], result["presentation"])
    return result


# =====================================================================
# Screens
# =====================================================================

def login_screen():
    inject_css()
    with st.form("login_form"):
        st.markdown(
            f'''<div class="login-heading">
                <img class="login-logo" src="{LOGO_URL}" />
                <div class="login-title">TNS Retail Intelligence</div>
                <div class="login-sub">Sign in to access company analytics</div>
            </div>''',
            unsafe_allow_html=True,
        )
        username = st.text_input("Username", autocomplete="username")
        password = st.text_input("Password", type="password", autocomplete="current-password")
        submitted = wide(st.form_submit_button, "Sign in")
    if submitted:
        if check_credentials(username.strip(), password):
            st.session_state.authenticated = True
            st.session_state.username = username.strip()
            st.session_state.active_chat_id = None
            st.rerun()
        else:
            st.error("Invalid username or password.")


def load_history(session_id: str):
    history = store.get_history(session_id, st.session_state.username)
    if history is None:
        st.error("Chat session not found.")
        return None
    return history


def render_sidebar():
    with st.sidebar:
        st.markdown(
            f'<div class="tns-brand"><img src="{LOGO_URL}"/><div class="tns-brand-name">TNS Retail Intelligence</div></div>',
            unsafe_allow_html=True,
        )
        if wide(st.button, "＋  New Chat", key="new_chat"):
            st.session_state.active_chat_id = None
            st.session_state.pending_prompt = ""
            st.rerun()
        st.markdown('<div class="section-label">Conversations</div>', unsafe_allow_html=True)
        sessions = store.list_sessions(st.session_state.username)
        if not sessions:
            st.caption("No conversations yet.")
        active = st.session_state.get("active_chat_id")
        for session in sessions:
            label = session["title"] or "New Chat"
            c1, c2 = st.columns([0.84, 0.16], gap="small")
            with c1:
                is_active = session["session_id"] == active
                if wide(
                    st.button, label,
                    key=f"chat_{session['session_id']}",
                    type="primary" if is_active else "secondary",
                ):
                    st.session_state.active_chat_id = session["session_id"]
                    st.session_state.pending_prompt = ""
                    st.rerun()
            with c2:
                if st.button("✕", key=f"delete_{session['session_id']}", help="Delete conversation"):
                    store.delete_session(session["session_id"], st.session_state.username)
                    if st.session_state.get("active_chat_id") == session["session_id"]:
                        st.session_state.active_chat_id = None
                    st.rerun()
        st.divider()
        st.caption(f"Signed in as {st.session_state.username}")
        if wide(st.button, "Logout", key="logout"):
            st.session_state.clear()
            st.rerun()


def main_app():
    inject_css()
    ensure_store()
    render_sidebar()

    active_id = st.session_state.get("active_chat_id")
    pending = st.session_state.pop("pending_prompt", "")
    typed = st.chat_input("Ask a question...", max_chars=5000)
    prompt = pending or typed

    history = load_history(active_id) if active_id else None
    title = history["title"] if history else "New Chat"

    st.markdown(
        f'''<div class="tns-header"><div class="tns-header-title">{html.escape(title)}</div><div class="tns-header-subtitle">Business Intelligence Assistant</div></div>''',
        unsafe_allow_html=True,
    )

    if history:
        conversation_id = store.get_conversation_id(active_id, st.session_state.username)
        messages = history["messages"]
        for index, message in enumerate(messages):
            render_message(message, conversation_id, f"{active_id}_{index}", is_last=(index == len(messages) - 1))
    elif not prompt:
        st.markdown(
            '''<div class="welcome"><h1>How can I help you?</h1><p>Ask questions about company sales, products, stores, brands and more.</p></div>''',
            unsafe_allow_html=True,
        )

    if prompt:
        with st.chat_message("user"):
            st.markdown(prompt.replace("$", "\\$"))
        succeeded = False
        with st.chat_message("assistant"):
            with st.spinner("TNS is analyzing your request..."):
                try:
                    if active_id:
                        process_followup(active_id, prompt)
                    else:
                        process_new_message(prompt)
                    succeeded = True
                except Exception as exc:
                    logger.exception("TNS request failed")
                    st.error(f"Request failed: {exc}")
        if succeeded:
            st.rerun()


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