import asyncio
import base64
import csv
import hashlib
import hmac
import html
import importlib.util
import io
import logging
import math
import re
import json
import time
from datetime import datetime
from typing import Any, Dict, List, Optional

import httpx
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

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
    st.markdown(
        """
        <style>
        #MainMenu {visibility:hidden;}
        footer {visibility:hidden;}
        header {background:transparent !important;}
        .stApp {background:#f5f7fb; color:#172033;}

        /* ---------- layout ---------- */
        [data-testid="stMainBlockContainer"], .block-container {
            max-width: 980px !important;
            padding-top: 1.2rem !important;
            padding-bottom: 7rem !important;
        }

        /* ---------- sidebar ---------- */
        [data-testid="stSidebar"] {background:#111827; min-width:270px; max-width:270px;}
        [data-testid="stSidebar"] * {color:#d1d5db;}
        [data-testid="stSidebar"] .stButton button {
            justify-content:flex-start; text-align:left;
            border:1px solid #374151; background:#1f2937; color:white;
            border-radius:9px; min-height:38px;
        }
        [data-testid="stSidebar"] .stButton button p {
            overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:13px; color:#e5e7eb;
        }
        [data-testid="stSidebar"] .stButton button:hover {background:#374151; border-color:#4b5563;}
        [data-testid="stSidebar"] .stButton button[kind="primary"],
        [data-testid="stSidebar"] .stButton button[data-testid="stBaseButton-primary"] {
            background:#374151; border-color:#6b7280;
        }
        .tns-brand {display:flex;align-items:center;gap:10px;padding:4px 0 18px;}
        .tns-brand img {width:38px;height:38px;border-radius:9px;object-fit:contain;background:white;}
        .tns-brand-name {font-weight:650;font-size:16px;color:white !important;}
        .section-label {font-size:11px;font-weight:650;text-transform:uppercase;letter-spacing:.7px;color:#9ca3af;margin:14px 0 7px;}

        /* ---------- header / welcome ---------- */
        .tns-header {border-bottom:1px solid #e5e7eb; padding:2px 0 14px; margin-bottom:20px;}
        .tns-header-title {font-weight:600;font-size:16px;color:#172033;}
        .tns-header-subtitle {font-size:11px;color:#9ca3af;margin-top:2px;}
        .welcome {text-align:center;margin:14vh auto 8vh;}
        .welcome h1 {font-size:30px;letter-spacing:-.5px;color:#172033;margin-bottom:10px;}
        .welcome p {font-size:14px;color:#6b7280;}

        /* ---------- chat bubbles ---------- */
        [data-testid="stChatMessage"] {background:transparent; padding:0; gap:0; margin-bottom:16px;}
        [data-testid^="stChatMessageAvatar"] {display:none !important;}
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {
            background:#111827; border-radius:14px 14px 4px 14px;
            padding:12px 16px; margin-left:20%;
        }
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) * {color:#ffffff !important;}
        [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) {
            background:#ffffff; border:1px solid #e5e7eb;
            border-radius:14px 14px 14px 4px; padding:14px 18px; color:#1f2937;
        }
        [data-testid="stChatMessage"] h1 {font-size:20px; margin:14px 0 6px;}
        [data-testid="stChatMessage"] h2 {font-size:17px; margin:14px 0 6px;}
        [data-testid="stChatMessage"] h3 {font-size:15px; margin:14px 0 6px;}
        [data-testid="stChatMessage"] h4 {font-size:14px; margin:12px 0 6px;}
        [data-testid="stChatMessage"] p, [data-testid="stChatMessage"] li {font-size:14px; line-height:1.6;}
        [data-testid="stChatMessage"] table {border-collapse:collapse; font-size:13px; margin:8px 0 10px;}
        [data-testid="stChatMessage"] th, [data-testid="stChatMessage"] td {
            border:1px solid #e5e7eb !important; padding:6px 10px !important; text-align:left;
        }
        [data-testid="stChatMessage"] th {background:#f9fafb !important; font-weight:700;}
        [data-testid="stChatMessage"] sup {font-size:10px; color:#2563eb; font-weight:700; padding:0 1px;}

        /* ---------- cards, buttons ---------- */
        [data-testid="stChatMessage"] [data-testid="stVerticalBlockBorderWrapper"] {
            border-radius:12px; background:#ffffff;
        }
        [data-testid="stMain"] .stButton button,
        [data-testid="stMain"] .stDownloadButton button,
        section.main .stButton button,
        section.main .stDownloadButton button {
            border:1px solid #d1d5db; background:#ffffff; color:#374151;
            border-radius:999px; font-size:12px; padding:3px 12px; min-height:32px;
        }
        [data-testid="stMain"] .stButton button:hover,
        [data-testid="stMain"] .stDownloadButton button:hover {background:#f1f5f9; border-color:#cbd5e1; color:#111827;}
        [data-testid="stMain"] .stButton button p, [data-testid="stMain"] .stDownloadButton button p {font-size:12px;}

        /* ---------- login ---------- */
        [data-testid="stForm"] {
            max-width:430px !important; margin:12vh auto 0 !important; padding:32px !important;
            background:#ffffff !important; border:1px solid #e5e7eb !important;
            border-radius:16px !important; box-shadow:0 15px 40px rgba(0,0,0,.07) !important;
        }
        [data-testid="stForm"] [data-testid="stTextInput"] label,
        [data-testid="stForm"] [data-testid="stTextInput"] label p {
            color:#374151 !important; font-size:13px !important; font-weight:600 !important;
        }
        [data-testid="stForm"] [data-testid="stTextInput"] input {
            color:#111827 !important; -webkit-text-fill-color:#111827 !important;
            background:#ffffff !important; border:1px solid #d1d5db !important; border-radius:9px !important;
        }
        [data-testid="stForm"] button {
            color:#ffffff !important; -webkit-text-fill-color:#ffffff !important;
            background:#111827 !important; border:1px solid #111827 !important;
            border-radius:9px !important; font-weight:600 !important;
        }
        .login-logo {display:block;width:72px;height:72px;object-fit:contain;margin:0 auto 14px;border-radius:14px;background:#ffffff;border:1px solid #e5e7eb;box-shadow:0 2px 8px rgba(0,0,0,.06);}
        .login-title{text-align:center;font-size:24px;font-weight:700;color:#172033;margin:0 0 7px;}
        .login-sub{text-align:center;color:#6b7280;font-size:14px;margin:0 0 28px;}
        .login-heading {text-align:center;}

        /* ---------- chat composer ---------- */
        [data-testid="stBottom"], [data-testid="stBottom"] > div {background:#f5f7fb !important;}
        [data-testid="stChatInput"] {
            background:#ffffff !important; border:1px solid #d1d5db !important;
            border-radius:14px !important; box-shadow:0 4px 15px rgba(0,0,0,.05) !important;
        }
        [data-testid="stChatInput"] textarea, [data-testid="stChatInput"] input,
        [data-testid="stChatInput"] textarea:focus, [data-testid="stChatInput"] input:focus {
            color:#111827 !important; -webkit-text-fill-color:#111827 !important;
            caret-color:#111827 !important; background:#ffffff !important;
        }
        [data-testid="stChatInput"] textarea::placeholder, [data-testid="stChatInput"] input::placeholder {
            color:#6b7280 !important; -webkit-text-fill-color:#6b7280 !important; opacity:1 !important;
        }

        /* =====================================================
           FORCE A READABLE LIGHT UI
           Streamlit follows the visitor's system theme. In dark mode it paints
           text near-white, which is invisible on our white cards. Everything
           below pins colours explicitly so the app reads correctly either way.
           ===================================================== */
        .stApp {color-scheme: light;}
        .stApp [data-testid="stMain"], .stApp section.main {color:#1f2937;}

        .stApp [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"])
            :is(p, li, span, div, h1, h2, h3, h4, h5, h6, td, th, label, strong, em, b, i, summary, small) {
            color:#1f2937 !important;
        }
        .stApp [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"])
            :is(p, li, span, div, h1, h2, h3, h4, h5, h6, strong, em, a, code) {
            color:#ffffff !important;
        }
        .stApp.stApp [data-testid="stChatMessage"] [data-testid="stCaptionContainer"],
        .stApp.stApp [data-testid="stChatMessage"] [data-testid="stCaptionContainer"] * {color:#6b7280 !important;}
        .stApp.stApp [data-testid="stChatMessage"] sup {color:#2563eb !important; font-weight:700;}
        .stApp.stApp [data-testid="stChatMessage"] a {color:#2563eb !important;}
        .stApp.stApp [data-testid="stChatMessage"] .section-label {color:#9ca3af !important;}
        .stApp.stApp [data-testid="stChatMessage"] th {background:#f9fafb !important;}
        .stApp.stApp [data-testid="stChatMessage"] table {background:#ffffff !important;}
        .stApp.stApp [data-testid="stChatMessage"] code,
        .stApp.stApp [data-testid="stChatMessage"] [data-testid="stCode"],
        .stApp.stApp [data-testid="stChatMessage"] [data-testid="stCode"] pre,
        .stApp.stApp [data-testid="stChatMessage"] [data-testid="stCode"] * {
            background:#f3f4f6 !important; color:#111827 !important;
        }

        /* cards / expanders */
        .stApp [data-testid="stChatMessage"] [data-testid="stVerticalBlockBorderWrapper"] {
            background:#ffffff !important; border-color:#e5e7eb !important;
        }
        .stApp [data-testid="stExpander"], .stApp [data-testid="stExpander"] details,
        .stApp [data-testid="stExpander"] summary {
            background:#ffffff !important; border-color:#e5e7eb !important;
        }
        .stApp [data-testid="stExpander"] summary:hover {background:#f9fafb !important;}

        /* pill buttons in the main area (downloads, suggestions, full CSV) */
        .stApp [data-testid="stMain"] :is(.stButton, .stDownloadButton, [data-testid="stDownloadButton"]) button,
        .stApp section.main :is(.stButton, .stDownloadButton) button {
            background:#ffffff !important; color:#374151 !important;
            border:1px solid #d1d5db !important; border-radius:999px !important;
        }
        .stApp [data-testid="stMain"] :is(.stButton, .stDownloadButton, [data-testid="stDownloadButton"]) button:hover {
            background:#f1f5f9 !important; border-color:#9ca3af !important;
        }
        .stApp [data-testid="stMain"] :is(.stButton, .stDownloadButton, [data-testid="stDownloadButton"]) button * {
            color:#374151 !important;
        }

        /* chat composer: the dark rounded bar came from an inner wrapper */
        .stApp [data-testid="stChatInput"],
        .stApp [data-testid="stChatInput"] div,
        .stApp [data-testid="stChatInput"] textarea {
            background:#ffffff !important; color:#111827 !important;
            -webkit-text-fill-color:#111827 !important; caret-color:#111827 !important;
        }
        .stApp [data-testid="stChatInput"] {border:1px solid #d1d5db !important; border-radius:14px !important;}
        .stApp [data-testid="stChatInput"] div {border-color:transparent !important;}
        .stApp [data-testid="stChatInput"] textarea::placeholder {
            color:#6b7280 !important; -webkit-text-fill-color:#6b7280 !important; opacity:1 !important;
        }
        .stApp [data-testid="stChatInput"] button {background:#111827 !important;}
        .stApp [data-testid="stChatInput"] button svg {fill:#ffffff !important; color:#ffffff !important;}
        .stApp [data-testid="stBottom"], .stApp [data-testid="stBottom"] > div {background:#f5f7fb !important;}
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


def _auth_secret() -> bytes:
    secret = getattr(settings, "app_auth_secret", None) or settings.app_password
    return str(secret).encode("utf-8")


def _make_auth_token(username: str) -> str:
    payload = {"u": username, "exp": int(time.time()) + 30 * 24 * 3600}
    raw = base64.urlsafe_b64encode(
        json.dumps(payload, separators=(",", ":")).encode("utf-8")
    ).decode("ascii").rstrip("=")
    sig = hmac.new(_auth_secret(), raw.encode("ascii"), hashlib.sha256).hexdigest()
    return f"{raw}.{sig}"


def _read_auth_token() -> Optional[str]:
    token = st.query_params.get("tns_auth")
    if not token or "." not in token:
        return None
    raw, sig = token.rsplit(".", 1)
    expected = hmac.new(_auth_secret(), raw.encode("ascii"), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(sig, expected):
        return None
    try:
        padded = raw + "=" * (-len(raw) % 4)
        payload = json.loads(base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8"))
        username = str(payload.get("u") or "")
        if username != settings.app_username or int(payload.get("exp", 0)) < int(time.time()):
            return None
        return username
    except Exception:
        return None


def set_authenticated(username: str) -> None:
    st.session_state.authenticated = True
    st.session_state.username = username
    st.query_params["tns_auth"] = _make_auth_token(username)


def clear_authenticated() -> None:
    st.session_state.clear()
    try:
        del st.query_params["tns_auth"]
    except Exception:
        pass


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
# Charts — FastAPI-compatible Chart.js renderer
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
        if any(re.search(rf"\b{re.escape(w)}s?\b", title_text) for w in words)
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
        "firstKind": first_kind,
        "horizontal": chart_type == "bar" and len(rows) > 12,
    }


def _chart_js_html(spec: Dict[str, Any], title: str, source_title: str) -> str:
    payload = json.dumps(spec, ensure_ascii=False, separators=(",", ":"))
    title_json = json.dumps(title or "Visualization", ensure_ascii=False)
    source_json = json.dumps(source_title or "Query result", ensure_ascii=False)
    colors_json = json.dumps(CHART_COLORS)
    template = """<!doctype html>
<html><head><meta charset="utf-8">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
html,body{margin:0;padding:0;background:#fff;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#374151}
.title{font-size:14px;font-weight:650;color:#172033;padding:4px 14px 8px}
.chart-wrap{position:relative;width:100%;padding:12px 14px 4px;box-sizing:border-box;background:#fff}
.source{padding:4px 14px 10px;color:#6b7280;font-size:11px}
.error{padding:16px;color:#b91c1c;font-size:13px}
</style></head><body>
<div class="title" id="title"></div><div class="chart-wrap" id="chartWrap"><canvas id="chart"></canvas></div><div class="source" id="source"></div>
<script>
const spec=__SPEC__;
const title=__TITLE__;
const source=__SOURCE__;
const COLORS=__COLORS__;
document.getElementById("title").textContent=title;
document.getElementById("source").textContent="Source: "+source;
function formatValue(value, kind) {
  if (value === null || value === undefined || value === "") return "—";
  const n=Number(value);
  if (!Number.isFinite(n)) return String(value);
  if (kind === "percent") return n.toLocaleString("en-IN",{maximumFractionDigits:2})+"%";
  if (kind === "currency") return "₹"+n.toLocaleString("en-IN",{maximumFractionDigits:0});
  return n.toLocaleString("en-IN",{maximumFractionDigits:2});
}
try {
  if (typeof Chart === "undefined") throw new Error("Chart.js failed to load");
  const wrap=document.getElementById("chartWrap");
  wrap.style.height=spec.horizontal ? Math.max(340,spec.labels.length*24+90)+"px" : "340px";
  const valueScale={beginAtZero:true,ticks:{callback:(value)=>formatValue(value,spec.firstKind)}};
  new Chart(document.getElementById("chart"),{
    type:spec.type,
    data:{labels:spec.labels,datasets:spec.datasets.map((dataset,index)=>({
      label:dataset.label,data:dataset.data,hidden:dataset.hidden,
      backgroundColor:COLORS[index%COLORS.length],borderColor:COLORS[index%COLORS.length],
      borderWidth:spec.type==="line"?2:0,borderRadius:spec.type==="bar"?3:0,
      pointRadius:spec.type==="line"?3:0,tension:0.25,spanGaps:false
    }))},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:spec.horizontal?"y":"x",
      plugins:{legend:{display:spec.datasets.length>1},tooltip:{callbacks:{label:(context)=>{
        const value=spec.horizontal?context.parsed.x:context.parsed.y;
        const kind=spec.datasets[context.datasetIndex].kind;
        return context.dataset.label+": "+formatValue(value,kind);
      }}}},
      scales:spec.horizontal?{x:valueScale,y:{ticks:{autoSkip:false}}}:{y:valueScale,x:{ticks:{maxRotation:60,autoSkip:spec.labels.length>24}}}
    }
  });
} catch (e) {
  document.getElementById("chartWrap").innerHTML='<div class="error">Unable to render chart: '+String(e.message||e)+'</div>';
}
</script></body></html>"""
    return (template.replace("__SPEC__", payload)
                    .replace("__TITLE__", title_json)
                    .replace("__SOURCE__", source_json)
                    .replace("__COLORS__", colors_json))


@st.cache_data(show_spinner=False, ttl=3600, max_entries=32)
def fetch_visualization_png(conversation_id: str, message_id: str, attachment_id: str) -> bytes:
    return call_genie(lambda client: client.download_visualization(conversation_id, message_id, attachment_id))


def render_visualization(viz: Dict[str, Any], tables: List[Dict[str, Any]], conversation_id: str,
                         message_id: str, key: str):
    with st.container(border=True):
        title = viz.get("title") or "Visualization"
        table = find_source_table(viz, tables)
        if table:
            try:
                spec = build_chart_spec(table, viz)
                if spec:
                    components.html(
                        _chart_js_html(spec, title, table.get("title") or "Query result"),
                        height=(max(410, len(spec["labels"])*24+145) if spec["horizontal"] else 410),
                        scrolling=False,
                    )
                    with st.expander("View data"):
                        render_table(table, conversation_id, message_id, f"{key}_data")
                    return
            except Exception as exc:
                logger.warning("FastAPI-compatible Chart.js rendering failed: %s", exc)

        attachment_id = viz.get("attachment_id")
        if attachment_id and message_id:
            try:
                image = fetch_visualization_png(conversation_id, message_id, attachment_id)
                wide(st.image, image)
                st.download_button("⬇ Chart (PNG)", data=image, file_name="TNS_visualization.png",
                                   mime="image/png", key=f"viz_dl_{key}")
                return
            except Exception as exc:
                logger.warning("Visualization retrieval failed: %s", exc)
        st.info("This visualization is not available.")


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


def _presentation_visualizations(presentation: Dict[str, Any]) -> List[Dict[str, Any]]:
    if not isinstance(presentation, dict):
        return []
    visualizations = presentation.get("visualizations") or []
    if visualizations:
        return [v for v in visualizations if isinstance(v, dict)]
    for block in presentation.get("blocks") or []:
        if isinstance(block, dict) and block.get("type") == "visualization" and isinstance(block.get("data"), dict):
            visualizations.append(block["data"])
    return visualizations


def _presentation_tables(presentation: Dict[str, Any]) -> List[Dict[str, Any]]:
    if not isinstance(presentation, dict):
        return []
    tables = presentation.get("tables") or []
    if tables:
        return [t for t in tables if isinstance(t, dict)]
    return [
        block["data"]
        for block in (presentation.get("blocks") or [])
        if isinstance(block, dict) and block.get("type") == "table" and isinstance(block.get("data"), dict)
    ]


def _needs_previous_chart(prompt: str, answer: str) -> bool:
    text = f"{prompt} {answer}".lower()
    return bool(re.search(r"\b(chart|charts|plot|plots|graph|graphs|visuali[sz]ation|visuali[sz]ations)\b", text))


def recover_previous_visualization(session_id: str, prompt: str, answer: str,
                                   current_presentation: Dict[str, Any]) -> Dict[str, Any]:
    """FastAPI-compatible history behavior for follow-ups such as 'give me the chart'.

    Genie can answer a follow-up by referring to a visualization already attached
    to an earlier message while returning attachments=[] on the new message. The
    FastAPI UI still has the earlier presentation in the conversation DOM/history.
    Reuse that exact stored visualization + source query table instead of inventing
    a new chart or trying to download the current message's nonexistent attachment.
    """
    if _presentation_visualizations(current_presentation):
        return current_presentation
    if not _needs_previous_chart(prompt, answer):
        return current_presentation

    history = store.get_history(session_id, st.session_state.username)
    if not history:
        return current_presentation

    messages = history.get("messages") or []
    for previous in reversed(messages):
        if previous.get("role") != "assistant":
            continue
        presentation = previous.get("metadata") or previous.get("presentation") or {}
        previous_viz = _presentation_visualizations(presentation)
        if not previous_viz:
            continue

        previous_tables = _presentation_tables(presentation)
        recovered = dict(current_presentation or {})
        recovered["visualizations"] = list(previous_viz)

        current_tables = _presentation_tables(recovered)
        existing_ids = {str(t.get("attachment_id")) for t in current_tables if t.get("attachment_id")}
        for viz in previous_viz:
            source_id = viz.get("query_attachment_id") or viz.get("attachment_id")
            for table in previous_tables:
                if source_id and table.get("attachment_id") == source_id and str(source_id) not in existing_ids:
                    current_tables.append(table)
                    existing_ids.add(str(source_id))

        recovered["tables"] = current_tables
        recovered["blocks"] = [
            {"type": "visualization", "data": viz}
            for viz in previous_viz
        ]
        for table in current_tables:
            recovered["blocks"].append({"type": "table", "data": table})

        if presentation.get("agent_message_id"):
            recovered["agent_message_id"] = presentation["agent_message_id"]

        logger.info(
            "Recovered %d previous Genie visualization(s) from chat history for follow-up: %s",
            len(previous_viz), prompt,
        )
        return recovered

    return current_presentation


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
    result["presentation"] = recover_previous_visualization(
        session_id, prompt, result["answer"], result["presentation"]
    )
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
            set_authenticated(username.strip())
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
            clear_authenticated()
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
    restored_user = _read_auth_token()
    if restored_user:
        st.session_state.authenticated = True
        st.session_state.username = restored_user
    else:
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