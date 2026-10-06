import asyncio
import html
import logging
from typing import Any, Dict, Optional, List, Tuple

import httpx
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

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
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("tns_streamlit")


def inject_css():
    st.markdown(
        """
        <style>
        #MainMenu, footer {visibility:hidden;}
        header {background:transparent !important;}

        :root {
            --tns-navy:#111827;
            --tns-navy-2:#1f2937;
            --tns-text:#172033;
            --tns-muted:#667085;
            --tns-border:#e4e7ec;
            --tns-bg:#f6f8fb;
            --tns-card:#ffffff;
            --tns-accent:#2563eb;
        }

        .stApp {background:var(--tns-bg); color:var(--tns-text);}
        .main .block-container {
            max-width:1180px;
            padding-top:1.15rem;
            padding-bottom:7rem;
            padding-left:1.4rem;
            padding-right:1.4rem;
        }

        /* Sidebar */
        [data-testid="stSidebar"] {
            background:var(--tns-navy);
            min-width:300px;
            max-width:300px;
            border-right:1px solid #202938;
        }
        [data-testid="stSidebar"] > div:first-child {padding-top:1.1rem;}
        [data-testid="stSidebar"] * {color:#e5e7eb;}
        [data-testid="stSidebar"] .stButton button {
            min-height:42px;
            border:1px solid #344054;
            background:#1d2939;
            color:#f9fafb !important;
            border-radius:10px;
            box-shadow:none;
            font-weight:500;
        }
        [data-testid="stSidebar"] .stButton button:hover {
            background:#273449;
            border-color:#52617a;
        }
        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {gap:.35rem;}

        .tns-brand {
            display:flex;
            align-items:center;
            gap:11px;
            padding:3px 0 20px;
        }
        .tns-brand img {
            width:44px;
            height:44px;
            border-radius:10px;
            object-fit:contain;
            background:white;
        }
        .tns-brand-name {font-weight:700;font-size:16px;color:#fff;}
        .sidebar-section {
            margin:12px 0 8px;
            color:#98a2b3 !important;
            font-size:11px;
            font-weight:700;
            letter-spacing:.08em;
            text-transform:uppercase;
        }
        .sidebar-user {
            border-top:1px solid #344054;
            margin-top:18px;
            padding-top:18px;
            color:#98a2b3 !important;
            font-size:13px;
        }

        /* Header */
        .tns-header {
            display:flex;
            align-items:center;
            justify-content:space-between;
            min-height:64px;
            background:#fff;
            border:1px solid var(--tns-border);
            border-radius:14px;
            padding:12px 20px;
            margin:0 0 22px;
            box-shadow:0 1px 2px rgba(16,24,40,.03);
        }
        .tns-header-title {font-weight:700;font-size:16px;color:var(--tns-text);}
        .tns-header-subtitle {font-size:12px;color:#98a2b3;margin-top:3px;}

        /* Welcome */
        .welcome {
            text-align:center;
            padding:15vh 0 20vh;
        }
        .welcome h1 {
            font-size:34px;
            line-height:1.15;
            letter-spacing:-.8px;
            color:var(--tns-text);
            margin:0 0 12px;
        }
        .welcome p {font-size:14px;color:var(--tns-muted);margin:0;}

        /* Conversation layout */
        .message-shell {
            width:100%;
            margin:0 0 20px;
        }
        .user-row {
            display:flex;
            justify-content:flex-end;
            width:100%;
        }
        .assistant-row {
            display:flex;
            justify-content:flex-start;
            width:100%;
        }
        .user-bubble {
            display:inline-block;
            max-width:74%;
            background:var(--tns-navy);
            color:#fff;
            border-radius:16px 16px 5px 16px;
            padding:12px 16px;
            line-height:1.55;
            font-size:14px;
            overflow-wrap:anywhere;
            box-shadow:0 2px 6px rgba(16,24,40,.08);
        }
        .assistant-bubble {
            width:min(100%, 1040px);
            background:#fff;
            border:1px solid var(--tns-border);
            border-radius:16px 16px 16px 5px;
            padding:18px 20px;
            line-height:1.65;
            font-size:14px;
            color:var(--tns-text);
            box-shadow:0 2px 8px rgba(16,24,40,.035);
        }
        .assistant-bubble p:first-child {margin-top:0;}
        .assistant-bubble p:last-child {margin-bottom:0;}
        .assistant-answer {margin-bottom:15px;}

        /* Result cards */
        .result-card-head {
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:12px;
            margin-bottom:10px;
        }
        .result-title {font-weight:700;color:var(--tns-text);font-size:15px;}
        .result-subtitle {font-size:12px;color:var(--tns-muted);margin-top:3px;}
        .result-divider {height:1px;background:#eef0f3;margin:10px 0 14px;}
        .section-label {
            font-size:11px;
            font-weight:700;
            text-transform:uppercase;
            letter-spacing:.08em;
            color:#98a2b3;
            margin:16px 0 8px;
        }

        /* Streamlit widgets inside result cards */
        div[data-testid="stExpander"] {
            border:1px solid var(--tns-border) !important;
            border-radius:10px !important;
            background:#fff !important;
        }
        div[data-testid="stExpander"] summary {color:var(--tns-text) !important;}
        div[data-testid="stDataFrame"] {
            border:1px solid #eaecf0;
            border-radius:10px;
            overflow:hidden;
        }
        .stDownloadButton button, .stLinkButton a {
            border:1px solid #d0d5dd !important;
            background:#fff !important;
            color:#344054 !important;
            border-radius:8px !important;
            font-size:13px !important;
        }

        /* Suggestions */
        .suggestions-note {font-size:12px;color:#667085;margin:14px 0 8px;}
        
        /* Login */
        .login-page {
            min-height:78vh;
            display:flex;
            align-items:center;
            justify-content:center;
        }
        .login-card {
            width:min(440px, 100%);
            background:#fff;
            border:1px solid var(--tns-border);
            border-radius:18px;
            padding:36px 38px 30px;
            box-shadow:0 18px 45px rgba(16,24,40,.08);
        }
        .login-logo {display:block;width:58px;height:58px;object-fit:contain;margin:0 auto 16px;border-radius:12px;background:#fff;}
        .login-title {text-align:center;font-size:26px;font-weight:750;color:var(--tns-text);}
        .login-sub {text-align:center;color:var(--tns-muted);font-size:13px;margin:7px 0 25px;}
        .login-card label, .login-card [data-testid="stWidgetLabel"] p {
            color:#344054 !important;
            font-weight:600 !important;
            font-size:13px !important;
        }
        .login-card input {
            color:#172033 !important;
            background:#fff !important;
            caret-color:#172033 !important;
        }
        .login-card input::placeholder {color:#98a2b3 !important;}
        .login-card .stButton button {
            background:#111827 !important;
            color:#fff !important;
            border:1px solid #111827 !important;
            border-radius:9px !important;
            min-height:44px;
            font-weight:650;
        }

        /* Chat composer */
        div[data-testid="stChatInput"] {
            background:#fff !important;
            border:1px solid #d0d5dd !important;
            border-radius:15px !important;
            box-shadow:0 8px 25px rgba(16,24,40,.10) !important;
        }
        div[data-testid="stChatInput"] textarea {
            color:#172033 !important;
            background:#fff !important;
            caret-color:#172033 !important;
            font-size:15px !important;
        }
        div[data-testid="stChatInput"] textarea::placeholder {color:#98a2b3 !important;opacity:1 !important;}
        div[data-testid="stChatInput"] button {
            background:#111827 !important;
            color:#fff !important;
            border-radius:10px !important;
        }

        /* Plotly */
        .plotly-chart-wrap {
            border:1px solid #eaecf0;
            border-radius:12px;
            background:#fff;
            padding:4px 4px 0;
        }

        @media (max-width: 900px) {
            [data-testid="stSidebar"] {min-width:260px;max-width:260px;}
            .user-bubble {max-width:88%;}
            .assistant-bubble {width:100%;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# Every async operation creates, uses, and closes its client on ONE event loop.
def run_async(coro_factory):
    async def runner():
        return await coro_factory()
    return asyncio.run(runner())


async def _call_genie_async(coro_factory):
    client = GenieClient()
    try:
        return await coro_factory(client)
    finally:
        await client.close()


def call_genie(coro_factory):
    return run_async(lambda: _call_genie_async(coro_factory))


def check_credentials(username: str, password: str) -> bool:
    return username == settings.app_username and password == settings.app_password


def ensure_store():
    if st.session_state.get("store_initialized"):
        return
    store.initialize()
    st.session_state.store_initialized = True


def login_screen():
    inject_css()
    st.markdown(
        f'''<div class="login-page"><div class="login-card">
            <img class="login-logo" src="{LOGO_URL}" />
            <div class="login-title">TNS Retail Intelligence</div>
            <div class="login-sub">Sign in to access company analytics</div>
        </div></div>''',
        unsafe_allow_html=True,
    )
    # Keep the actual Streamlit form visually attached to the login card.
    with st.container():
        with st.form("login_form"):
            username = st.text_input("Username", autocomplete="username")
            password = st.text_input("Password", type="password", autocomplete="current-password")
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


def render_thoughts(thoughts):
    if not thoughts:
        return
    with st.expander("Thought process", expanded=False):
        for thought in thoughts:
            content = thought.get("content") if isinstance(thought, dict) else str(thought)
            if content:
                st.markdown(content)


def download_signed_links(table: Dict[str, Any], conversation_id: str, message_id: str):
    attachment_id = table.get("attachment_id")
    if not attachment_id or not message_id:
        return
    try:
        links = call_genie(lambda client: client.get_full_query_download_links(conversation_id, message_id, attachment_id))
        for index, link in enumerate(links):
            try:
                response = httpx.get(link, timeout=120.0)
                response.raise_for_status()
                label = "Download query result" if len(links) == 1 else f"Download query result {index + 1}"
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


def render_table(table: Dict[str, Any], conversation_id: str, message_id: str):
    with st.container(border=True):
        title = table.get("title") or "Query result"
        description = table.get("description")
        st.markdown(
            f'<div class="result-card-head"><div><div class="result-title">{html.escape(str(title))}</div>'
            + (f'<div class="result-subtitle">{html.escape(str(description))}</div>' if description else "")
            + '</div></div><div class="result-divider"></div>',
            unsafe_allow_html=True,
        )
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


def _pick_date_column(df: pd.DataFrame) -> Optional[str]:
    for col in df.columns:
        name = str(col).lower()
        if any(token in name for token in ("date", "week", "month", "day", "time")):
            parsed = pd.to_datetime(df[col], errors="coerce")
            if parsed.notna().sum() >= max(2, int(len(df) * 0.6)):
                return col
    return None


def _plotly_from_table(viz: Dict[str, Any], table: Dict[str, Any]) -> Optional[go.Figure]:
    """Create a polished Plotly chart only when the Genie query table is chartable.

    This is presentation-only. Genie remains the source of truth for the
    answer, query, result and native visualization attachment.
    """
    columns = table.get("columns") or []
    rows = table.get("rows") or []
    if len(columns) < 2 or not rows:
        return None
    df = pd.DataFrame(rows, columns=columns)
    if df.empty:
        return None

    numeric = list(df.select_dtypes(include="number").columns)
    if not numeric:
        return None

    x_col = _pick_date_column(df)
    if x_col is not None:
        parsed = pd.to_datetime(df[x_col], errors="coerce")
        df[x_col] = parsed
    else:
        non_numeric = [c for c in df.columns if c not in numeric]
        if not non_numeric:
            return None
        x_col = non_numeric[0]

    # Avoid plotting IDs / row numbers when a better business metric exists.
    numeric = [c for c in numeric if not str(c).lower().endswith("_id")]
    if not numeric:
        return None
    y_cols = numeric[:4]

    title = str(viz.get("title") or table.get("title") or "Analysis")
    if len(y_cols) == 1:
        fig = px.line(
            df,
            x=x_col,
            y=y_cols[0],
            markers=True,
            title=title,
            template="plotly_white",
        )
    else:
        fig = px.line(
            df,
            x=x_col,
            y=y_cols,
            markers=True,
            title=title,
            template="plotly_white",
        )

    fig.update_layout(
        height=430,
        margin=dict(l=55, r=30, t=65, b=55),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(family="Inter, Segoe UI, sans-serif", color="#344054", size=12),
        title=dict(font=dict(size=17, color="#172033"), x=0.02, xanchor="left"),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0, xanchor="left"),
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, linecolor="#d0d5dd", tickfont=dict(color="#667085"))
    fig.update_yaxes(gridcolor="#eaecf0", zeroline=False, tickfont=dict(color="#667085"))
    fig.update_traces(line=dict(width=3), marker=dict(size=7))
    return fig


def render_visualization(viz: Dict[str, Any], conversation_id: str, message_id: str, tables: List[Dict[str, Any]]):
    with st.container(border=True):
        title = viz.get("title") or "Visualization"
        st.markdown(
            f'<div class="result-card-head"><div class="result-title">{html.escape(str(title))}</div></div>',
            unsafe_allow_html=True,
        )

        chart_table = None
        query_attachment_id = viz.get("query_attachment_id")
        if query_attachment_id:
            chart_table = next((t for t in tables if t.get("attachment_id") == query_attachment_id), None)

        # Prefer a polished Plotly rendering when the query result contains
        # enough structured data to reproduce the chart safely. Otherwise use
        # the exact native Genie PNG.
        fig = _plotly_from_table(viz, chart_table) if chart_table else None
        if fig is not None:
            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displaylogo": False, "responsive": True, "modeBarButtonsToRemove": ["lasso2d", "select2d"]},
            )
        else:
            attachment_id = viz.get("attachment_id")
            if attachment_id and message_id:
                try:
                    image = call_genie(lambda client: client.download_visualization(conversation_id, message_id, attachment_id))
                    st.image(image, use_container_width=True)
                    st.download_button(
                        "Download chart",
                        data=image,
                        file_name="genie_visualization.png",
                        mime="image/png",
                        key=f"viz_download_{conversation_id}_{message_id}_{attachment_id}",
                    )
                except Exception as exc:
                    logger.warning("Visualization retrieval failed: %s", exc)
                    st.warning("The analytical response was returned, but this visualization could not be loaded.")


def render_presentation(presentation: Dict[str, Any], conversation_id: str, message_id: str):
    if not presentation:
        return

    blocks = presentation.get("blocks") or []
    tables = presentation.get("tables") or []
    visualizations = presentation.get("visualizations") or []

    # Preserve the exact Genie block order. The Streamlit layer only changes
    # presentation styling; it does not reorder the response.
    if blocks:
        rendered_table_ids = set()
        rendered_viz_ids = set()
        for block in blocks:
            block_type = block.get("type")
            data = block.get("data")
            if block_type == "thoughts":
                render_thoughts(data or [])
            elif block_type == "table" and data:
                render_table(data, conversation_id, message_id)
                if data.get("attachment_id"):
                    rendered_table_ids.add(data.get("attachment_id"))
            elif block_type == "visualization" and data:
                render_visualization(data, conversation_id, message_id, tables)
                if data.get("attachment_id"):
                    rendered_viz_ids.add(data.get("attachment_id"))
            elif block_type == "suggested_questions" and data:
                st.markdown('<div class="suggestions-note">Suggested follow-up questions</div>', unsafe_allow_html=True)
                for index, question in enumerate(data):
                    if st.button(question, key=f"suggestion_{message_id}_{index}", use_container_width=True):
                        st.session_state.pending_prompt = question
                        st.rerun()

        # Compatibility for older saved presentations that have top-level
        # objects not represented in blocks.
        for table in tables:
            if table.get("attachment_id") not in rendered_table_ids:
                render_table(table, conversation_id, message_id)
        for viz in visualizations:
            if viz.get("attachment_id") not in rendered_viz_ids:
                render_visualization(viz, conversation_id, message_id, tables)
        return

    # Backward-compatible presentations from before ordered Agent blocks.
    render_thoughts(presentation.get("thoughts") or [])
    for viz in visualizations:
        render_visualization(viz, conversation_id, message_id, tables)
    for table in tables:
        render_table(table, conversation_id, message_id)
    suggestions = presentation.get("suggested_questions") or []
    if suggestions:
        st.markdown('<div class="suggestions-note">Suggested follow-up questions</div>', unsafe_allow_html=True)
        for index, question in enumerate(suggestions):
            if st.button(question, key=f"suggestion_legacy_{message_id}_{index}", use_container_width=True):
                st.session_state.pending_prompt = question
                st.rerun()


def render_message(message: Dict[str, Any], conversation_id: str):
    role = message.get("role")
    content = message.get("content") or ""
    if role == "user":
        safe = html.escape(content).replace("\n", "<br>")
        st.markdown(
            f'<div class="message-shell"><div class="user-row"><div class="user-bubble">{safe}</div></div></div>',
            unsafe_allow_html=True,
        )
        return

    st.markdown('<div class="message-shell"><div class="assistant-row"><div class="assistant-bubble">', unsafe_allow_html=True)
    st.markdown(content)
    st.markdown('</div></div></div>', unsafe_allow_html=True)
    presentation = message.get("presentation") or message.get("metadata") or {}
    message_id = presentation.get("agent_message_id") or message.get("message_id") or ""
    render_presentation(presentation, conversation_id, message_id)


async def run_agent_turn(client: GenieClient, message: str, conversation_id: Optional[str]):
    original = conversation_id
    rebound = False
    try:
        response = await client.create_agent_response(message, conversation_id=conversation_id, enable_visualization=True)
    except GenieError as exc:
        if conversation_id and client.is_legacy_conversation_error(exc):
            rebound = True
            response = await client.create_agent_response(message, conversation_id=None, enable_visualization=True)
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


async def _process_new_message_async(prompt: str):
    client = GenieClient()
    try:
        return await run_agent_turn(client, prompt, None)
    finally:
        await client.close()


def process_new_message(prompt: str):
    result = run_async(lambda: _process_new_message_async(prompt))
    title = prompt.strip()
    if len(title) > 60:
        title = title[:57] + "..."
    sid = store.create_session(result["conversation_id"], st.session_state.username, title)
    store.add_message(sid, st.session_state.username, "user", prompt)
    store.add_message(sid, st.session_state.username, "assistant", result["answer"], result["presentation"])
    st.session_state.active_chat_id = sid
    return result


async def _process_followup_async(session_id: str, prompt: str, conversation_id: str):
    client = GenieClient()
    try:
        return await run_agent_turn(client, prompt, conversation_id)
    finally:
        await client.close()


def process_followup(session_id: str, prompt: str):
    conversation_id = store.get_conversation_id(session_id, st.session_state.username)
    if not conversation_id:
        raise GenieError("Chat session not found.")
    result = run_async(lambda: _process_followup_async(session_id, prompt, conversation_id))
    if result["conversation_changed"]:
        store.set_conversation_id(session_id, st.session_state.username, result["conversation_id"])
    store.add_message(session_id, st.session_state.username, "user", prompt)
    store.add_message(session_id, st.session_state.username, "assistant", result["answer"], result["presentation"])
    return result


def render_sidebar():
    with st.sidebar:
        st.markdown(
            f'<div class="tns-brand"><img src="{LOGO_URL}"/><div class="tns-brand-name">TNS Retail Intelligence</div></div>',
            unsafe_allow_html=True,
        )
        if st.button("＋  New Chat", use_container_width=True):
            st.session_state.active_chat_id = None
            st.session_state.pending_prompt = ""
            st.rerun()
        st.markdown('<div class="sidebar-section">Conversations</div>', unsafe_allow_html=True)
        sessions = get_sessions()
        for session in sessions:
            label = session["title"] or "New Chat"
            c1, c2 = st.columns([0.86, 0.14], gap="small")
            with c1:
                if st.button(label, key=f"chat_{session['session_id']}", use_container_width=True):
                    st.session_state.active_chat_id = session["session_id"]
                    st.session_state.pending_prompt = ""
                    st.rerun()
            with c2:
                if st.button("×", key=f"delete_{session['session_id']}", help="Delete conversation"):
                    store.delete_session(session["session_id"], st.session_state.username)
                    if st.session_state.get("active_chat_id") == session["session_id"]:
                        st.session_state.active_chat_id = None
                    st.rerun()
        st.markdown('<div class="sidebar-user">Signed in as <strong>' + html.escape(st.session_state.username) + '</strong></div>', unsafe_allow_html=True)
        if st.button("Logout", use_container_width=True):
            st.session_state.clear()
            st.rerun()


def main_app():
    inject_css()
    ensure_store()
    render_sidebar()

    active_id = st.session_state.get("active_chat_id")
    history = load_history(active_id) if active_id else None
    title = history["title"] if history else "New Chat"

    st.markdown(
        f'''<div class="tns-header"><div><div class="tns-header-title">{html.escape(str(title))}</div><div class="tns-header-subtitle">Business Intelligence Assistant</div></div></div>''',
        unsafe_allow_html=True,
    )

    if history:
        conversation_id = store.get_conversation_id(active_id, st.session_state.username)
        for message in history["messages"]:
            render_message(message, conversation_id)
    else:
        st.markdown(
            '''<div class="welcome"><h1>How can I help you?</h1><p>Ask questions about company sales, products, stores, brands and more.</p></div>''',
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
