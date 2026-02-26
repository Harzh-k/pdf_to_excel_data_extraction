"""
app.py  —  IRDAI PDF Extractor  |  Streamlit UI
Run:   streamlit run app.py
"""

import os
import time
import logging
import tempfile
from pathlib import Path
import  base64
import streamlit as st

def get_base64_image(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
image_path = os.path.join(BASE_DIR, "pdf", "bajaj-life-logo.png")

img_base64 = get_base64_image(image_path)
LOGO_IMAGE = "https://la.bajajlife.com/assets/Icons/General/BalicLogo2.png"
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Drag-n-Fly",
    page_icon=LOGO_IMAGE,
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# THEME CONFIG  — force light mode so dark-theme users get consistent UI
# ─────────────────────────────────────────────────────────────────────────────
# Write a .streamlit/config.toml if it doesn't exist yet
_cfg_dir = Path(".streamlit")
_cfg_file = _cfg_dir / "config.toml"
if not _cfg_file.exists():
    _cfg_dir.mkdir(exist_ok=True)
    _cfg_file.write_text(
        "[theme]\nbase=\"light\"\nprimaryColor=\"#1A56DB\"\n"
        "backgroundColor=\"#F0F4F8\"\nsecondaryBackgroundColor=\"#FFFFFF\"\n"
        "textColor=\"#111827\"\nfont=\"sans serif\"\n"
    )

# ─────────────────────────────────────────────────────────────────────────────
# CSS  — targets real Streamlit DOM nodes, no wrapping-div tricks
# ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
# <style>
#
# /* ── Globals ── */
# html, body { font-family: 'Inter', sans-serif !important; }
# #MainMenu, footer, header { visibility: hidden; }
# .stApp { background: #F0F4F8 !important; }
# .block-container {
#     max-width: 680px !important;
#     padding: 2rem 1.5rem 4rem !important;
# }
#
# /* ── All text dark by default ── */
# .stApp, .stApp p, .stApp label, .stApp span,
# .stApp div, .stApp h1, .stApp h2, .stApp h3 {
#     color: #111827;
# }
#
# /* ── Streamlit containers become cards ── */
# [data-testid="stVerticalBlockBorderWrapper"] > div {
#     background: #FFFFFF !important;
#     border: 1px solid #E2E8F0 !important;
#     border-radius: 14px !important;
#     box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important;
#     padding: 4px 0 !important;
# }
#
# /* ── File uploader — force light style ── */
# [data-testid="stFileUploader"] {
#     background: transparent !important;
# }
# [data-testid="stFileUploaderDropzone"],
# [data-testid="stFileUploader"] > div,
# [data-testid="stFileUploader"] section {
#     background: #F8FAFF !important;
#     border: 2px dashed #BFDBFE !important;
#     border-radius: 10px !important;
#     color: #374151 !important;
# }
# [data-testid="stFileUploaderDropzoneInstructions"] {
#     color: #374151 !important;
# }
# [data-testid="stFileUploaderDropzoneInstructions"] span,
# [data-testid="stFileUploaderDropzoneInstructions"] small,
# [data-testid="stFileUploaderDropzoneInstructions"] p {
#     color: #6B7280 !important;
# }
# /* Browse files button */
# [data-testid="stFileUploader"] button,
# [data-testid="stBaseButton-secondary"] {
#     background: #EFF6FF !important;
#     color: #1D4ED8 !important;
#     border: 1px solid #BFDBFE !important;
#     border-radius: 7px !important;
#     font-weight: 600 !important;
# }
#
# /* ── Text input ── */
# .stTextInput input {
#     background: #FFFFFF !important;
#     border: 1.5px solid #D1D5DB !important;
#     border-radius: 8px !important;
#     color: #111827 !important;
#     font-size: 0.9rem !important;
#     padding: 10px 12px !important;
# }
# .stTextInput input:focus {
#     border-color: #1A56DB !important;
#     box-shadow: 0 0 0 3px rgba(26,86,219,0.1) !important;
# }
# .stTextInput label {
#     color: #6B7280 !important;
#     font-size: 0.75rem !important;
#     font-weight: 600 !important;
#     text-transform: uppercase !important;
#     letter-spacing: 0.07em !important;
# }
#
# /* ── Buttons ── */
# .stButton > button {
#     border-radius: 9px !important;
#     font-weight: 600 !important;
#     font-size: 0.9rem !important;
#     padding: 11px 20px !important;
#     transition: all 0.15s !important;
# }
# .stButton > button[kind="primary"] {
#     background: #1A56DB !important;
#     color: #fff !important;
#     border: none !important;
#     box-shadow: 0 2px 8px rgba(26,86,219,0.3) !important;
# }
# .stButton > button[kind="primary"]:hover {
#     background: #1447C0 !important;
# }
# .stButton > button:disabled {
#     opacity: 0.45 !important;
# }
#
# /* ── Download button ── */
# .stDownloadButton > button {
#     background: #16A34A !important;
#     color: #fff !important;
#     border: none !important;
#     border-radius: 9px !important;
#     font-weight: 600 !important;
#     font-size: 0.9rem !important;
#     padding: 11px 20px !important;
#     box-shadow: 0 2px 8px rgba(22,163,74,0.3) !important;
#     width: 100% !important;
# }
# .stDownloadButton > button:hover {
#     background: #15803D !important;
# }
#
# /* ── Progress bar ── */
# .stProgress > div > div > div > div {
#     background: linear-gradient(90deg, #1A56DB, #60A5FA) !important;
#     border-radius: 4px !important;
# }
#
# /* ── Alert / info boxes ── */
# [data-testid="stAlert"] {
#     border-radius: 10px !important;
# }
#
# /* ── Metric numbers ── */
# [data-testid="stMetric"] {
#     background: #F8FAFF !important;
#     border: 1px solid #E2E8F0 !important;
#     border-radius: 12px !important;
#     padding: 16px 12px !important;
#     text-align: center !important;
# }
# [data-testid="stMetric"] label {
#     color: #6B7280 !important;
#     font-size: 0.7rem !important;
#     font-weight: 600 !important;
#     text-transform: uppercase !important;
#     letter-spacing: 0.08em !important;
#     justify-content: center !important;
# }
# [data-testid="stMetricValue"] {
#     color: #1A56DB !important;
#     font-size: 1.8rem !important;
#     font-weight: 700 !important;
#     justify-content: center !important;
# }
# [data-testid="stMetricDelta"] { display: none !important; }
#
# /* ── Divider ── */
# hr { border: none !important; border-top: 1px solid #E5E7EB !important; margin: 8px 0 !important; }
#
# /* ── Caption / small text ── */
# .stApp .stCaption, small { color: #6B7280 !important; }
#
# /* ── Sidebar (hidden but clean) ── */
# [data-testid="stSidebar"] { display: none; }
#
# </style>
# """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────

class _UILogHandler(logging.Handler):
    def emit(self, record):
        if "logs" not in st.session_state:
            st.session_state["logs"] = []
        try:
            msg = self.format(record)
        except Exception:
            msg = str(record.getMessage())
        st.session_state["logs"].append((record.levelname, msg))


def _setup_logging():
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    if not any(isinstance(h, _UILogHandler) for h in root.handlers):
        h = _UILogHandler()
        h.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
        root.addHandler(h)


def _log_html() -> str:
    logs = st.session_state.get("logs", [])
    if not logs:
        return ""
    css = {"INFO": "#93C5FD", "WARNING": "#FCD34D", "ERROR": "#FCA5A5"}
    lines = []
    for lvl, msg in logs[-60:]:
        color = css.get(lvl, "#6EE7B7")
        safe  = msg.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        lines.append(f'<span style="color:{color}">{safe}</span>')
    return (
        '<div style="background:#0F172A;border-radius:9px;padding:12px 14px;'
        'max-height:180px;overflow-y:auto;font-family:Courier New,monospace;'
        'font-size:0.71rem;line-height:1.75;">'
        + "<br>".join(lines) + "</div>"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Extraction  — identical to CLI  (extract_document + write_excel)
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _run_extraction(pdf_bytes: bytes):
    from extract_tables_smart_merged import extract_document, write_excel

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name

    try:
        document = extract_document(tmp_path)

        pages_seen, tables, forms = set(), 0, set()
        for pb in document:
            pages_seen.add(pb.get("page_number", 0))
            if pb.get("form_name"):
                forms.add(pb["form_name"])
            for item in pb.get("content", []):
                if item.get("type") == "table":
                    tables += 1

        excel_bytes = write_excel(document)   # returns raw bytes

        stats = {
            "pages":  max(pages_seen) if pages_seen else 0,
            "tables": tables,
            "sheets": len(document),
            "forms":  sorted(forms),
        }
        return excel_bytes, stats

    finally:
        os.unlink(tmp_path)


# ─────────────────────────────────────────────────────────────────────────────
# UI — uses only native Streamlit widgets, NO div-wrapping tricks
# ─────────────────────────────────────────────────────────────────────────────

def main():
    _setup_logging()

    # ── Page title ────────────────────────────────────────────────────────────
    # Adjusted column ratios to give the logo more breathing room
    col_logo, col_title, col_badge = st.columns([3, 6,2], gap="large", vertical_alignment="center")

    with col_logo:
        st.markdown(
            f'''
            <div style="display: flex; align-items: center; justify-content: center;">
                <img src="data:image/png;base64,{img_base64}" 
                     style="width: 180px; height: auto; object-fit: contain; transform: scale(1.1);">
            </div>
            ''',
            unsafe_allow_html=True,
        )

    with col_title:
        st.markdown(
            '''
            <div style="display: flex; flex-direction: column; justify-content: center;">
                <h1 style="margin: 0; font-size: 1.8rem; font-weight: 100; color: #0F172A; line-height: 1;">
                    Powered by <span style="font-weight: 400; color: #64748B;">Drag-n-Fly</span>
                </h1>
                <p style="margin: 4px 0 0; font-size: 0.95rem; color: #64748B; font-weight: 400; max-width: 300px;">
                    Converting public disclosure PDFs to structured Excel
                </p>
            </div>
            ''',
            unsafe_allow_html=True,
        )

    with col_badge:
        st.markdown(
            '''
            <div style="display: flex; justify-content: flex-end;">
                <span style="
                    background: #F1F5F9; 
                    border: 1px solid #E2E8F0; 
                    color: #475569;
                    font-size: 0.65rem; 
                    font-weight: 700; 
                    padding: 5px 10px; 
                    border-radius: 4px;
                    text-transform: uppercase;
                    letter-spacing: 0.05em;">
                    FORM L-SERIES • V2.0
                </span>
            </div>
            ''',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Upload section ────────────────────────────────────────────────────────
    st.markdown(
        '<p style="font-size:0.72rem;font-weight:600;letter-spacing:0.09em;'
        'text-transform:uppercase;color:#9CA3AF;margin:0 0 8px">📁 Upload PDF</p>',
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Upload IRDAI PDF",
        type=["pdf"],
        help="Any IRDAI public disclosure PDF — max 200 MB",
        label_visibility="collapsed",
    )

    if uploaded:
        mb = uploaded.size / 1_048_576
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:12px;background:#EFF6FF;'
            f'border:1px solid #BFDBFE;border-radius:10px;padding:10px 14px;margin-top:8px">'
            f'<span style="font-size:1.4rem">📄</span>'
            f'<div><p style="font-weight:600;font-size:0.87rem;color:#1E40AF;margin:0">'
            f'{uploaded.name}</p>'
            f'<p style="font-size:0.72rem;color:#64748B;margin:2px 0 0">'
            f'{mb:.2f} MB · PDF ready</p></div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Filename + Button row ─────────────────────────────────────────────────
    col_name, col_btn = st.columns([3, 2], gap="medium")

    with col_name:
        st.markdown(
            '<p style="font-size:0.72rem;font-weight:600;letter-spacing:0.09em;'
            'text-transform:uppercase;color:#9CA3AF;margin:0 0 6px">⚙️ Output filename</p>',
            unsafe_allow_html=True,
        )
        stem = Path(uploaded.name).stem if uploaded else "output"
        output_name = st.text_input(
            "Output filename",
            value=f"{stem}_IRDAI.xlsx",
            label_visibility="collapsed",
        )
        if not output_name.lower().endswith(".xlsx"):
            output_name += ".xlsx"

    with col_btn:
        st.markdown(
            '<p style="font-size:0.72rem;font-weight:600;letter-spacing:0.09em;'
            'text-transform:uppercase;color:#9CA3AF;margin:0 0 6px">⚡ Action</p>',
            unsafe_allow_html=True,
        )
        extract_btn = st.button(
            "⚡  Extract to Excel",
            type="primary",
            use_container_width=True,
            disabled=(uploaded is None),
        )

    # ── Progress ──────────────────────────────────────────────────────────────
    prog_slot = st.empty()
    msg_slot  = st.empty()

    # ── Run ───────────────────────────────────────────────────────────────────
    if extract_btn and uploaded:
        for k in ("excel_bytes", "stats", "logs"):
            st.session_state.pop(k, None)

        prog = prog_slot.progress(0, text="Reading PDF…")
        msg_slot.info("Processing — may take a moment for large files…", icon="⏳")

        try:
            pdf_bytes = uploaded.read()
            prog.progress(20, text="Detecting page types…")
            excel_bytes, stats = _run_extraction(pdf_bytes)
            prog.progress(95, text="Building Excel…")
            time.sleep(0.25)
            prog.progress(100, text="Done ✓")
            time.sleep(0.3)
            prog_slot.empty()
            msg_slot.empty()
            st.session_state["excel_bytes"] = excel_bytes
            st.session_state["stats"]       = stats
            st.rerun()

        except Exception as exc:
            prog_slot.empty()
            msg_slot.error(f"Extraction failed: {exc}", icon="🚨")
            import traceback
            st.session_state.setdefault("logs", [])
            st.session_state["logs"].append(("ERROR", traceback.format_exc()))

    # ── Results ───────────────────────────────────────────────────────────────
    excel_bytes = st.session_state.get("excel_bytes")
    stats       = st.session_state.get("stats", {})

    if excel_bytes and stats:
        st.markdown("---")

        # Success banner
        p = stats.get("pages",  0)
        t = stats.get("tables", 0)
        s = stats.get("sheets", 0)
        fname = uploaded.name if uploaded else "document"
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:12px;background:#F0FDF4;'
            f'border:1px solid #BBF7D0;border-radius:10px;padding:13px 16px;margin-bottom:16px">'
            f'<span style="font-size:1.5rem;flex-shrink:0">✅</span>'
            f'<div><p style="font-weight:600;font-size:0.9rem;color:#14532D;margin:0">'
            f'Extraction complete — {t} tables across {s} sheets</p>'
            f'<p style="font-size:0.73rem;color:#16A34A;margin:2px 0 0">'
            f'{fname} → {output_name}</p></div></div>',
            unsafe_allow_html=True,
        )

        # Stats — use st.metric so they render properly
        c1, c2, c3 = st.columns(3)
        c1.metric("Pages",  p)
        c2.metric("Tables", t)
        c3.metric("Sheets", s)

        # Form badges
        forms = stats.get("forms", [])
        if forms:
            badge_html = " ".join(
                f'<span style="background:#F0FDF4;border:1px solid #BBF7D0;color:#166534;'
                f'font-size:0.68rem;font-weight:600;padding:3px 10px;border-radius:6px;'
                f'display:inline-block;margin:3px 3px 3px 0">{f}</span>'
                for f in forms
            )
            st.markdown(f'<div style="margin:12px 0 4px">{badge_html}</div>',
                        unsafe_allow_html=True)

        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

        # Download
        st.download_button(
            label="⬇️  Download Excel",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    elif not uploaded:
        st.markdown("---")
        st.markdown(
            '<div style="text-align:center;padding:32px 0 16px">'
            '<div style="font-size:2.4rem;margin-bottom:10px;opacity:0.4">📂</div>'
            '<p style="font-weight:600;font-size:0.95rem;color:#475569;margin:0 0 6px">'
            'Upload a PDF to get started</p>'
            '<p style="font-size:0.78rem;color:#94A3B8;margin:0;line-height:1.8">'
            'Supports HDFC · Tata AIA · ICICI · LIC · Bajaj · Canara · Shriram<br>'
            'All IRDAI filers · FORM L-1-A-RA through L-26</p></div>',
            unsafe_allow_html=True,
        )

    # ── Log ───────────────────────────────────────────────────────────────────
    lh = _log_html()
    if lh:
        st.markdown("---")
        st.markdown(
            '<p style="font-size:0.72rem;font-weight:600;letter-spacing:0.09em;'
            'text-transform:uppercase;color:#9CA3AF;margin:0 0 8px">📋 Processing log</p>',
            unsafe_allow_html=True,
        )
        st.markdown(lh, unsafe_allow_html=True)


if __name__ == "__main__":
    main()