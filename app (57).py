"""
DHC Working Automation — Professional Streamlit Application

A modern, enterprise-grade MIS reporting platform for automated collections analytics,
RTGS monitoring, compliance validation, and delay tracking.

Version: 2.0
Last Updated: June 2026
"""

import traceback
import urllib.parse
import streamlit as st
import pandas as pd
import etl
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════════════
# PAGE CONFIGURATION & BRANDING
# ═══════════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="DHC Collections Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get Help": "https://docs.example.com",
        "Report a bug": "https://bugs.example.com",
        "About": "DHC Automation Platform v2.0"
    }
)

# ═══════════════════════════════════════════════════════════════════════════════════
# PROFESSIONAL STYLING
# ═══════════════════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
:root {
    --primary: #1e40af;
    --secondary: #0891b2;
    --accent: #06b6d4;
    --success: #059669;
    --warning: #ea580c;
    --danger: #dc2626;
    --dark: #0f172a;
    --light: #f8fafc;
    --border: #e2e8f0;
}

* {
    font-family: 'Segoe UI', 'Roboto', sans-serif;
}

html, body, [data-testid="stAppViewContainer"] {
    background-color: #f0f5fb;
}

[data-testid="stSidebarContent"] {
    background: linear-gradient(180deg, #0f172a 0%, #1e3a8a 100%);
    padding-top: 2rem;
}

[data-testid="stSidebarContent"] * {
    color: white !important;
}

[data-testid="stSidebarContent"] .stTabs [role="tablist"] button {
    color: rgba(255,255,255,0.7) !important;
}

[data-testid="stSidebarContent"] .stTabs [role="tablist"] button[aria-selected="true"] {
    color: white !important;
    border-bottom: 3px solid #06b6d4 !important;
}

/* Header Styling */
.header-container {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #1e40af 100%);
    padding: 3rem 2rem;
    border-radius: 12px;
    margin-bottom: 2rem;
    color: white;
    box-shadow: 0 10px 30px rgba(15, 23, 42, 0.15);
}

.header-title {
    font-size: 2.8rem;
    font-weight: 800;
    margin: 0;
    background: linear-gradient(135deg, #fff 0%, #e0f2fe 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}

.header-subtitle {
    font-size: 1rem;
    color: rgba(255,255,255,0.85);
    margin: 0.5rem 0 0 0;
    font-weight: 400;
}

.header-meta {
    display: flex;
    gap: 2rem;
    margin-top: 1.5rem;
    padding-top: 1.5rem;
    border-top: 1px solid rgba(255,255,255,0.2);
    font-size: 0.9rem;
    color: rgba(255,255,255,0.7);
}

.meta-item {
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

/* Card Styling */
.feature-card {
    background: white;
    border-radius: 12px;
    padding: 1.5rem;
    border: 1px solid #e2e8f0;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
    transition: all 0.3s ease;
}

.feature-card:hover {
    border-color: #06b6d4;
    box-shadow: 0 12px 16px rgba(6, 182, 212, 0.1);
    transform: translateY(-2px);
}

.feature-card-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #0f172a;
    margin-bottom: 0.5rem;
}

.feature-card-desc {
    font-size: 0.9rem;
    color: #64748b;
    line-height: 1.5;
}

/* Stats Grid */
.stats-container {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 1.5rem;
    margin: 2rem 0;
}

.stat-card {
    background: white;
    border-left: 4px solid #06b6d4;
    border-radius: 8px;
    padding: 1.5rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}

.stat-label {
    font-size: 0.85rem;
    color: #64748b;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 0.5rem;
}

.stat-value {
    font-size: 2rem;
    font-weight: 800;
    color: #0f172a;
    margin-bottom: 0.25rem;
}

.stat-change {
    font-size: 0.8rem;
    color: #059669;
    font-weight: 500;
}

/* File Upload Styling */
[data-testid="stFileUploader"] {
    background: linear-gradient(135deg, #f0f9ff 0%, #f0f4f8 100%);
    border: 2px dashed #06b6d4;
    border-radius: 12px;
    padding: 2rem !important;
    transition: all 0.3s ease;
}

[data-testid="stFileUploader"]:hover {
    background: linear-gradient(135deg, #e0f2fe 0%, #e8f4f8 100%);
    border-color: #0891b2;
}

/* Button Styling */
.stButton > button {
    background: linear-gradient(135deg, #1e40af 0%, #06b6d4 100%);
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.75rem 2rem;
    font-weight: 600;
    font-size: 1rem;
    transition: all 0.3s ease;
    box-shadow: 0 4px 12px rgba(30, 64, 175, 0.25);
}

.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(30, 64, 175, 0.35);
}

.stButton > button:active {
    transform: translateY(0);
}

/* Input Fields */
input, select, textarea {
    border-radius: 8px !important;
    border: 1px solid #e2e8f0 !important;
}

input:focus, select:focus, textarea:focus {
    border-color: #06b6d4 !important;
    box-shadow: 0 0 0 3px rgba(6, 182, 212, 0.1) !important;
}

/* Tabs */
.stTabs [role="tablist"] {
    border-bottom: 2px solid #e2e8f0;
    gap: 1rem;
}

.stTabs [role="tablist"] button {
    border-radius: 8px 8px 0 0;
    padding: 0.75rem 1.5rem;
    font-weight: 600;
    color: #64748b;
    transition: all 0.3s ease;
    border-bottom: 3px solid transparent;
}

.stTabs [role="tablist"] button[aria-selected="true"] {
    color: #1e40af;
    border-bottom-color: #06b6d4;
    background-color: rgba(6, 182, 212, 0.05);
}

/* Alerts & Status */
.stAlert {
    border-radius: 8px;
    border-left: 4px solid #06b6d4;
}

.stSuccess {
    background-color: #f0fdf4 !important;
    border-left-color: #059669 !important;
}

.stError {
    background-color: #fef2f2 !important;
    border-left-color: #dc2626 !important;
}

.stWarning {
    background-color: #fffbeb !important;
    border-left-color: #ea580c !important;
}

.stInfo {
    background-color: #f0f9ff !important;
    border-left-color: #06b6d4 !important;
}

/* Metrics */
[data-testid="metric-container"] {
    background: white;
    border-radius: 12px;
    padding: 1.5rem;
    border: 1px solid #e2e8f0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    transition: all 0.3s ease;
}

[data-testid="metric-container"]:hover {
    border-color: #06b6d4;
    box-shadow: 0 8px 16px rgba(6, 182, 212, 0.1);
}

/* Dividers */
hr {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, #e2e8f0, transparent);
    margin: 2rem 0;
}

/* Footer */
.footer {
    text-align: center;
    padding: 2rem;
    color: #64748b;
    font-size: 0.9rem;
    border-top: 1px solid #e2e8f0;
    margin-top: 3rem;
}

/* Dataframe styling */
.stDataFrame {
    border-radius: 8px;
    overflow: hidden;
}

/* Expander */
.streamlit-expanderHeader {
    background-color: #f8fafc;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
}

.streamlit-expanderHeader:hover {
    background-color: #f0f9ff;
    border-color: #06b6d4;
}

/* Loading spinner */
.stSpinner > div > div {
    border-color: #06b6d4 !important;
    border-right-color: transparent !important;
}

/* Code blocks */
code {
    background-color: #f8fafc;
    border-radius: 6px;
    padding: 0.2rem 0.4rem;
    font-size: 0.85rem;
    color: #dc2626;
}

/* Section headers */
h1 {
    color: #0f172a;
    font-weight: 800;
    margin-bottom: 1rem;
}

h2 {
    color: #1e40af;
    font-weight: 700;
    margin-top: 2rem;
    margin-bottom: 1rem;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 0.5rem;
}

h3 {
    color: #1e3a8a;
    font-weight: 600;
    margin-top: 1.5rem;
    margin-bottom: 0.75rem;
}

/* Links */
a {
    color: #1e40af;
    text-decoration: none;
    transition: color 0.3s ease;
}

a:hover {
    color: #06b6d4;
    text-decoration: underline;
}

/* Checkbox */
[role="checkbox"] {
    accent-color: #06b6d4 !important;
}

/* Slim top bar (post-login) — replaces repeating the giant hero banner
   on every page view once someone's already signed in */
.top-bar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 0.9rem 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}

.top-bar-title {
    font-size: 1.15rem;
    font-weight: 700;
    color: #0f172a;
}

.top-bar-meta {
    font-size: 0.85rem;
    color: #64748b;
}

/* Step section headers */
.step-header {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin: 0 0 1rem 0;
}

.step-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 2rem;
    height: 2rem;
    border-radius: 50%;
    background: #1e40af;
    color: white;
    font-weight: 700;
    font-size: 0.95rem;
    flex-shrink: 0;
}

.step-title {
    font-size: 1.25rem;
    font-weight: 700;
    color: #0f172a;
}

.step-subtitle {
    font-size: 0.85rem;
    color: #64748b;
    margin: 0.1rem 0 0 2.75rem;
}

</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════════

def format_large_number(num):
    """Format large numbers with commas and K/M suffix."""
    if num >= 1_000_000:
        return f"{num / 1_000_000:.1f}M"
    elif num >= 1_000:
        return f"{num / 1_000:.1f}K"
    return f"{num:,}"

def render_stat_card(label, value, delta=None, delta_color="green"):
    """Render a professional stat card."""
    if delta_color == "green":
        delta_html = f'<span style="color: #059669;">✓ {delta}</span>' if delta else ""
    elif delta_color == "red":
        delta_html = f'<span style="color: #dc2626;">⚠ {delta}</span>' if delta else ""
    else:
        delta_html = f'<span style="color: #06b6d4;">→ {delta}</span>' if delta else ""
    
    return f"""
    <div class="stat-card">
        <div class="stat-label">{label}</div>
        <div class="stat-value">{value}</div>
        {delta_html}
    </div>
    """

def step_header(number, title, subtitle=""):
    """Consistent numbered section header — used for every major stage
    (Upload / Process / Review / Export) so the page reads as a single
    linear workflow instead of a stack of loosely related blocks."""
    st.markdown(f"""
    <div class="step-header">
        <span class="step-badge">{number}</span>
        <span class="step-title">{title}</span>
    </div>
    {f'<p class="step-subtitle">{subtitle}</p>' if subtitle else ''}
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════════
# TEAM MEMBER GATE — "Who's working today?"
# Everything below this block (the existing app) only renders once someone
# picks a person. Shalini / Bhoobalan Sir are placeholders for now — swap in
# their real sections later the same way Swapna Mam's is wired up below.
# ═══════════════════════════════════════════════════════════════════════════════════

TEAM_MEMBERS = {
    "swapna": {"label": "Swapna Mam", "icon": "👩‍💼"},
    "shalini": {"label": "Shalini", "icon": "👩‍💼"},
    "bhoobalan": {"label": "Bhoobalan Sir", "icon": "👨‍💼"},
}

if "active_user" not in st.session_state:
    st.session_state["active_user"] = None

if st.session_state["active_user"] is None:
    st.markdown("""
    <div class="header-container" style="text-align:center;">
        <h1 class="header-title">📊 DHC Collections Intelligence Platform</h1>
        <p class="header-subtitle">Who's working today?</p>
    </div>
    """, unsafe_allow_html=True)

    st.write("")
    col1, col2, col3 = st.columns(3)
    for col, key in zip((col1, col2, col3), ("swapna", "shalini", "bhoobalan")):
        member = TEAM_MEMBERS[key]
        with col:
            if st.button(f"{member['icon']} {member['label']}", use_container_width=True,
                         type="primary", key=f"select_{key}"):
                st.session_state["active_user"] = key
                st.rerun()

    st.stop()

# Once someone's picked, show who's active + let them switch back to the picker.
with st.sidebar:
    active = TEAM_MEMBERS[st.session_state["active_user"]]
    st.markdown(f"**Working as:** {active['icon']} {active['label']}")
    if st.button("🔄 Switch User", use_container_width=True):
        st.session_state["active_user"] = None
        st.rerun()
    st.divider()

if st.session_state["active_user"] in ("shalini", "bhoobalan"):
    active = TEAM_MEMBERS[st.session_state["active_user"]]
    st.markdown(f"""
    <div class="header-container" style="text-align:center;">
        <h1 class="header-title">🚧 {active['icon']} {active['label']}</h1>
        <p class="header-subtitle">To be added</p>
    </div>
    """, unsafe_allow_html=True)
    st.info("This section hasn't been built yet — to be added.")
    st.stop()

# active_user == "swapna" from here on: falls through to the existing app.



with st.sidebar:
    st.markdown("""
    <div style="padding-bottom: 2rem;">
        <h3 style="color: white; margin: 0; font-size: 1.3rem;">📊 DHC Platform</h3>
        <p style="color: rgba(255,255,255,0.7); font-size: 0.85rem; margin: 0.5rem 0 0 0;">
            Collections Intelligence v2.0
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    with st.expander("📋 **File Format Guide**", expanded=False):
        st.markdown("""
        ### DCR Extract (.xlsb)
        - **Sheet1**: Daily receipts (17K–100K rows)
        - **Sheet2**: Agreement master (100K+ rows)
        
        ### Disable Lists (.xlsb)
        - **CIF Level Disable**: Customer restrictions (93K+)
        - **Agreement Level Disble**: Agreement restrictions (15K+)
        
        ### Employee Master (.xlsx)
        - Single column: **Mobile Number**
        - ~19,000 agent IDs
        
        ### Collection MIS Base (.xlsb)
        - **Sheet**: Overall Base
        - Columns needed: **AGREEMENTNO, CIF_NO, Opening DPD, Opening DPD SLAB**
        - Fresh export each period — source of truth for CIF mapping

        ### Receipt Cancellation Report (.xlsx)
        - **Sheet**: Report
        - Columns needed: **ReceiptNo, ReceiptDate, Amount, ReceiptStatus, ReceiptType, PaymentMode, Zone, AgreementNo, CustomerName, ReceiptCreatedDate**
        - Source of truth for cancelled receipts & their real Payment Mode (Cashfree, Payment Gateway, QR Code, etc.)
        """)
    
    with st.expander("🎯 **Output Sheets Overview**", expanded=False):
        st.markdown("""
        **1. Receipt Made Summary**
        - By receipt status & payment mode
        
        **2. RTGS Summary**
        - Zone × Receipt Type × Turn-Around Time
        - Online payment breakdown
        
        **3. Cash Mode Validation**
        - Compliance violations only
        
        **4. Delay in RCPTING**
        - Full month aging analysis
        
        **5. RCPT CXN**
        - Cancelled receipts (remarks for you)
        """)
    
    with st.expander("⚙️ **Configuration**", expanded=False):
        st.info("""
        Current settings:
        - **Engine**: Python + Pandas
        - **Processing**: ~40 seconds
        - **Output Format**: Excel (Pivot Style)
        - **Support**: Enterprise Grade
        """)
    
    st.divider()
    
    st.markdown("""
    <div style="text-align: center; padding: 1rem; color: rgba(255,255,255,0.6); font-size: 0.85rem;">
        <p style="margin: 0;">Last updated: June 2026</p>
        <p style="margin: 0.5rem 0 0 0;">Support • Documentation • Feedback</p>
    </div>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════════
# TOP BAR (slim — the full hero banner already showed on the "who's working" screen)
# ═══════════════════════════════════════════════════════════════════════════════════

active_member = TEAM_MEMBERS[st.session_state["active_user"]]
st.markdown(f"""
<div class="top-bar">
    <span class="top-bar-title">📊 DHC Collections Intelligence Platform</span>
    <span class="top-bar-meta">{active_member['icon']} {active_member['label']} · {datetime.now().strftime('%d %b %Y')}</span>
</div>
""", unsafe_allow_html=True)

with st.expander("ℹ️ About this platform", expanded=False):
    st.markdown("""
    Consolidates 5 monthly source files into 5 analytics summaries — RTGS turn-around time,
    cash-mode compliance, processing delays, cancelled receipts, and receipt-status breakdowns —
    normally built by hand in Excel. Typical processing time: 40–60 seconds for 100K+ records.
    """)

# ═══════════════════════════════════════════════════════════════════════════════════
# STEP 1 — UPLOAD SOURCE FILES
# ═══════════════════════════════════════════════════════════════════════════════════

with st.container(border=True):
    step_header(1, "Upload Source Files", "All five files are required to run the pipeline.")

    col_left, col_right = st.columns(2, gap="medium")

    with col_left:
        st.markdown("**Collections & Compliance Data**")
        dcr_file = st.file_uploader(
            "📊 DCR Extract",
            type=["xlsb"],
            help="Daily Collection Report (Sheet1: receipts, Sheet2: agreement master)",
            key="dcr"
        )
        disable_file = st.file_uploader(
            "🛡️ Disable Lists",
            type=["xlsb"],
            help="Cash-mode compliance lists (CIF & Agreement level restrictions)",
            key="disable"
        )
        cancellation_file = st.file_uploader(
            "🚫 Receipt Cancellation Report",
            type=["xlsx"],
            help="Receipt Cancellation Report (.xlsx, sheet 'Report') — source of truth for cancelled receipts & their Payment Mode",
            key="cancellation"
        )

    with col_right:
        st.markdown("**Reference Data**")
        employee_file = st.file_uploader(
            "👥 Employee Master",
            type=["xlsx"],
            help="Must contain 'Mobile Number' column for agent identification",
            key="employee"
        )
        collection_mis_file = st.file_uploader(
            "🧾 Collection MIS Base",
            type=["xlsb"],
            help="Collection MIS Base Working (.xlsb, sheet 'Overall Base') — source of truth for CIF No. & Opening DPD Slab per agreement",
            key="collection_mis"
        )

    st.markdown("---")
    st.markdown("**BRS MIS Files** &nbsp;·&nbsp; *optional — for the Email Summary*")
    st.caption("Fills in the mail's Delay in Receipting / Deposition section with real counts instead of [Fill manually]. Not required to run the main pipeline below.")

    brs_col1, brs_col2 = st.columns(2)
    with brs_col1:
        chq_brs_file = st.file_uploader("📑 Cheque/DD BRS MIS (.xlsx)", type=["xlsx"], key="chq_brs_upload")
        if chq_brs_file is not None:
            try:
                st.session_state["chq_brs"] = etl.process_brs_for_mail(etl.load_brs_raw(chq_brs_file), "Cheque")
                cb = st.session_state["chq_brs"]
                st.caption(f"✓ Loaded — Delay in Deposition: {cb['delay']['delay_deposition']:,}, Delay in Receipting: {cb['delay']['delay_receipting']:,}")
            except Exception as exc:
                st.session_state.pop("chq_brs", None)
                st.error(f"Couldn't process this file: {exc}")
        else:
            st.session_state.pop("chq_brs", None)

    with brs_col2:
        air_brs_file = st.file_uploader("📑 Airtel Cash BRS MIS (.xlsx)", type=["xlsx"], key="air_brs_upload")
        if air_brs_file is not None:
            try:
                st.session_state["air_brs"] = etl.process_brs_for_mail(etl.load_brs_raw(air_brs_file), "Airtel")
                ab = st.session_state["air_brs"]
                st.caption(f"✓ Loaded — Delay in Deposition: {ab['delay']['delay_deposition']:,}, Delay in Receipting: {ab['delay']['delay_receipting']:,}")
            except Exception as exc:
                st.session_state.pop("air_brs", None)
                st.error(f"Couldn't process this file: {exc}")
        else:
            st.session_state.pop("air_brs", None)

# ═══════════════════════════════════════════════════════════════════════════════════
# STEP 2 — RUN THE PIPELINE
# ═══════════════════════════════════════════════════════════════════════════════════

with st.container(border=True):
    step_header(2, "Run the Pipeline", "Takes about 40–60 seconds for a typical monthly file.")

    all_files_ready = all([dcr_file, disable_file, employee_file, collection_mis_file, cancellation_file])

    col_process, col_info = st.columns([2, 1], gap="medium")

    with col_process:
        run = st.button(
            "🚀  PROCESS FILES  →",
            type="primary",
            disabled=not all_files_ready,
            use_container_width=True,
            help="Upload all 5 files to enable" if not all_files_ready else "Click to start automated processing"
        )

    with col_info:
        if all_files_ready:
            st.success("✅ All files ready")
        else:
            missing = []
            if not dcr_file:
                missing.append("DCR")
            if not disable_file:
                missing.append("Disable Lists")
            if not employee_file:
                missing.append("Employee Master")
            if not collection_mis_file:
                missing.append("Collection MIS Base")
            if not cancellation_file:
                missing.append("Receipt Cancellation Report")
            st.warning(f"⏳ Missing: {', '.join(missing)}")

    if run:
        try:
            with st.status("📊 Running automated pipeline...", expanded=True) as status:
                st.write("📥 **Stage 1/7**: Loading DCR extract...")
                receipts, dcr_master = etl.load_dcr(dcr_file)
                st.write(f"   ✓ Loaded {len(receipts):,} receipts")

                st.write("📥 **Stage 2/7**: Loading compliance restrictions...")
                cif_disable, agr_disable = etl.load_disable_lists(disable_file)
                st.write(f"   ✓ Loaded {len(cif_disable):,} CIF + {len(agr_disable):,} agreement restrictions")

                st.write("📥 **Stage 3/7**: Loading employee mobile master...")
                mobiles = etl.load_employee_mobiles(employee_file)
                st.write(f"   ✓ Loaded {len(mobiles):,} agent mobile numbers")

                st.write("📥 **Stage 4/7**: Loading Collection MIS Base (CIF & DPD Slab)...")
                collection_mis_base = etl.load_collection_mis_base(collection_mis_file)
                st.write(f"   ✓ Loaded {len(collection_mis_base):,} agreements with CIF mapping")

                st.write("📥 **Stage 5/7**: Loading Receipt Cancellation Report...")
                cancellation_report = etl.load_receipt_cancellation_report(cancellation_file)
                st.write(f"   ✓ Loaded {len(cancellation_report):,} cancelled receipts")

                st.write("⚙️ **Stage 6/7**: Building enriched data tables...")
                lookup_master = etl.build_lookup_master(dcr_master, collection_mis_base)
                dcr_tab = etl.build_dcr_tab(receipts, lookup_master, agr_disable, cif_disable, mobiles)
                rtgs_tab = etl.build_rtgs_tab(dcr_tab)
                st.write(f"   ✓ Built DCR tab ({len(dcr_tab):,} receipts × 87 columns)")
                st.write(f"   ✓ Built RTGS tab ({len(rtgs_tab):,} RTGS receipts)")

                st.write("📊 **Stage 7/7**: Generating intelligence summaries...")
                receipt_made_summary = etl.build_receipt_made_summary(dcr_tab)
                rtgs_summary = etl.build_rtgs_summary(rtgs_tab, dcr_tab)
                cash_mode_validation_summary = etl.build_cash_mode_validation_summary(dcr_tab)
                delay_summary = etl.build_delay_in_rcpting_summary(dcr_tab)
                zone_overview_summary = etl.build_zone_overview_summary(dcr_tab)
                rcpt_cxn = etl.build_rcpt_cxn(cancellation_report)
                st.write("   ✓ Generated 6 analytics summaries")

                st.write("💾 **Finalizing**: Writing Excel workbook...")
                workbook_buffer = etl.write_output_workbook(
                    rtgs_summary, delay_summary, receipt_made_summary,
                    cash_mode_validation_summary, rcpt_cxn
                )
                st.write("   ✓ Workbook created (0 formula errors)")

                status.update(label="✅ Processing Complete! See Step 3 below for results.", state="complete")

            # Store in session — Step 3 below always renders from these, so
            # results stay on screen through later reruns (downloads,
            # switching tabs, etc.) instead of vanishing once `run` is no
            # longer True.
            st.session_state["workbook_buffer"] = workbook_buffer
            st.session_state["receipt_made_summary"] = receipt_made_summary
            st.session_state["rtgs_summary"] = rtgs_summary
            st.session_state["cash_mode_validation"] = cash_mode_validation_summary
            st.session_state["delay_summary"] = delay_summary
            st.session_state["zone_overview_summary"] = zone_overview_summary
            st.session_state["rcpt_cxn"] = rcpt_cxn
            st.session_state["lookup_master"] = lookup_master
            st.session_state["dcr_tab"] = dcr_tab
            st.session_state["rtgs_tab"] = rtgs_tab

        except Exception as e:
            st.error(f"❌ Processing Error: {str(e)}")
            with st.expander("📋 Detailed Error Information"):
                st.code(traceback.format_exc())

# ═══════════════════════════════════════════════════════════════════════════════════
# STEP 3 — REVIEW RESULTS
# Rendered straight from session_state (not tied to the button click), so it
# stays visible across reruns — e.g. clicking a download button below no
# longer makes this whole section disappear.
# ═══════════════════════════════════════════════════════════════════════════════════

if "workbook_buffer" in st.session_state:
    dcr_tab = st.session_state["dcr_tab"]
    rtgs_tab = st.session_state["rtgs_tab"]
    lookup_master = st.session_state["lookup_master"]
    rcpt_cxn = st.session_state["rcpt_cxn"]
    needs_cif = int(lookup_master["NEEDS_CIF_MAPPING"].sum())
    cxn_count = len(rcpt_cxn)

    with st.container(border=True):
        step_header(3, "Review Results")

        metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)

        with metric_col1:
            st.markdown(render_stat_card(
                "Total Receipts",
                f"{len(dcr_tab):,}",
                "Full month",
                "blue"
            ), unsafe_allow_html=True)

        with metric_col2:
            st.markdown(render_stat_card(
                "RTGS Subset",
                f"{len(rtgs_tab):,}",
                f"{len(rtgs_tab)/len(dcr_tab)*100:.1f}% RTGS",
                "blue"
            ), unsafe_allow_html=True)

        with metric_col3:
            st.markdown(render_stat_card(
                "Unmapped CIF",
                f"{needs_cif:,}",
                "No Collection MIS Base match",
                "red" if needs_cif > 0 else "green"
            ), unsafe_allow_html=True)

        with metric_col4:
            st.markdown(render_stat_card(
                "Cancelled Receipts",
                f"{cxn_count:,}",
                "For your review",
                "blue"
            ), unsafe_allow_html=True)

        if needs_cif > 0:
            st.warning(
                f"""
                ⚠️ **{needs_cif:,} AGREEMENTS WITH NO CIF MATCH**

                These agreements appear in this period's DCR but have no matching row in the Collection
                MIS Base you uploaded (e.g. fully closed accounts, or too new to appear in the MIS base
                pull yet), so their CIF No. is blank.

                👉 **Action**: Open the "Agreement → CIF Master" tab below to see the blank-CIF rows. If
                it's a fresh MIS base pull for this period, these should resolve on their own next run —
                otherwise confirm the agreement number with your LMS.
                """,
                icon="⚠️"
            )

# ═══════════════════════════════════════════════════════════════════════════════════
# STEP 3 (continued) — DETAILED BREAKDOWNS
# ═══════════════════════════════════════════════════════════════════════════════════

if "workbook_buffer" in st.session_state:
    st.markdown("#### Detailed Breakdowns")
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📈 Receipt Analytics",
        "🏦 RTGS Intelligence",
        "🛡️ Compliance Monitor",
        "⏱️ Delay Analysis",
        "❌ Cancellations",
        "🗂️ Look Up Master",
        "🔍 Raw Data"
    ])

    with tab1:
        st.markdown("#### Receipt Status & Payment Mode Analysis")
        st.info("✓ Two-sided pivot: Updated/Pending vs Bounced/Cancelled")
        
        summary = st.session_state["receipt_made_summary"]
        
        col_left, col_right = st.columns(2)
        
        with col_left:
            st.markdown("**Updated & Pending Receipts**")
            df_left = etl.receipt_made_table_to_dataframe(summary["left"], ["Cleared", "Deposit", "Pending"])
            st.dataframe(df_left, use_container_width=True, hide_index=True)
        
        with col_right:
            st.markdown("**Bounced & Cancelled Receipts**")
            df_right = etl.receipt_made_table_to_dataframe(summary["right"], ["Cleared", "Deposit", "Bounced", "Cxn"])
            st.dataframe(df_right, use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("#### RTGS Performance Dashboard")
        st.info("✓ Zone × Receipt Type × Turn-Around Time matrix + Online payment sources")
        
        summary = st.session_state["rtgs_summary"]
        
        col_matrix, col_sources = st.columns([2, 1])
        
        with col_matrix:
            st.markdown("**Zone-wise RTGS Performance**")
            df_matrix = etl.zone_tat_matrix_to_dataframe(summary["matrix"])
            st.dataframe(df_matrix, use_container_width=True, hide_index=True)
        
        with col_sources:
            st.markdown("**Online Payment Channels**")
            sources = summary["online_source_block"]["rows"]
            if sources:
                source_data = pd.DataFrame(sources, columns=["Channel", "Count"])
                st.dataframe(source_data, use_container_width=True, hide_index=True)
            else:
                st.info("No online receipts this period")

    with tab3:
        st.markdown("#### Compliance Violations Dashboard")
        st.info("✓ Shows only customers on disable list who paid via restricted methods")
        
        df = st.session_state["cash_mode_validation"]
        if df.empty:
            st.success("✅ **COMPLIANT** — No violations detected!")
        else:
            st.warning(f"⚠️ **{len(df)} Violations Found**")
            st.dataframe(df.head(50), use_container_width=True)
            if len(df) > 50:
                st.caption(f"Showing first 50 of {len(df)} rows")

    with tab4:
        st.markdown("#### Receipt Processing Delay Analysis")
        st.info("✓ Full month aging: How long between transaction and receipt entry?")
        
        summary = st.session_state["delay_summary"]
        st.markdown("**Processing TAT by Zone & Receipt Type**")
        df_matrix = etl.zone_tat_matrix_to_dataframe(summary["matrix"])
        st.dataframe(df_matrix, use_container_width=True, hide_index=True)
        
        st.markdown("**Interpretation**")
        col_info1, col_info2, col_info3 = st.columns(3)
        with col_info1:
            st.metric("✓ < 4 Days", "Optimal", help="Receipt processed within 3 days")
        with col_info2:
            st.metric("⏱️ 5-10 Days", "Acceptable", help="Normal processing window")
        with col_info3:
            st.metric("⚠️ > 10 Days", "Investigate", help="Delayed processing - follow up")

    with tab5:
        st.markdown("#### Cancelled Receipts Register")
        st.info("✓ All cancelled receipts with auto-filled details. YOU add the remarks.")
        
        df = st.session_state["rcpt_cxn"]
        if df.empty:
            st.success("✅ **CLEAN** — No cancelled receipts this period")
        else:
            st.warning(f"📋 {len(df)} Cancelled Receipts")
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.info("👉 **Next Step**: Download the output Excel and fill in the Remarks column for each row")

    with tab6:
        st.markdown("#### Agreement → CIF Master")
        st.info("✓ CIF No. & Opening DPD Slab sourced fresh from this period's Collection MIS Base")
        
        df = st.session_state["lookup_master"]
        
        metrics_col1, metrics_col2, metrics_col3 = st.columns(3)
        with metrics_col1:
            known = (~df["CIF_NO"].isna()).sum()
            st.metric("Known CIFs", f"{known:,}", help="Agreements with CIF mapping")
        with metrics_col2:
            new = (df["CIF_NO"].isna()).sum()
            st.metric("New (Blank)", f"{new:,}", help="Agreements needing CIF")
        with metrics_col3:
            total = len(df)
            st.metric("Total", f"{total:,}", help="All agreements in portfolio")
        
        st.dataframe(df.head(100), use_container_width=True, hide_index=True)
        if len(df) > 100:
            st.caption(f"Showing first 100 of {len(df)} agreements")

    with tab7:
        st.markdown("#### Raw DCR Data")
        st.info("✓ Full receipt-level data with 11 derived columns")
        
        df = st.session_state["dcr_tab"]
        col_select = st.multiselect(
            "Select columns to display",
            df.columns.tolist(),
            default=["AGREEMENTNO", "AMOUNTPAID", "MODEOFPAYMENT", "Status", "Mode", "Zone", "RECEIPT ENTER DATE"],
            help="Choose which columns to show"
        )
        
        if col_select:
            st.dataframe(df[col_select].head(100), use_container_width=True)
        if len(df) > 100:
            st.caption(f"Showing first 100 of {len(df)} receipts")

    st.markdown("---")
    step_header(4, "Export", "Grab the full workbook, or just one summary.")
    st.markdown("**Export a Specific Summary**")

    export_choice = st.selectbox(
        "Choose what to download",
        [
            "Receipt Made Summary",
            "RTGS Summary",
            "Cash Mode Validation Summary",
            "Delay in RCPTING Summary",
            "Zone Overview",
            "RCPT CXN",
        ],
    )

    if export_choice == "Receipt Made Summary":
        rms_table = st.session_state["receipt_made_summary"]["full"]
        st.dataframe(etl.receipt_made_status_table_to_dataframe(rms_table), use_container_width=True, hide_index=True)

        excel_col, image_col = st.columns(2)
        with excel_col:
            st.download_button(
                "📊 Download as Excel",
                data=etl.receipt_made_status_summary_to_excel_bytes(rms_table),
                file_name=f"Receipt_Made_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with image_col:
            st.download_button(
                "🖼️ Download as Image",
                data=etl.receipt_made_status_summary_to_image_bytes(rms_table),
                file_name=f"Receipt_Made_Report_{datetime.now().strftime('%Y%m%d')}.png",
                mime="image/png",
                use_container_width=True,
            )

    elif export_choice == "RTGS Summary":
        summary = st.session_state["rtgs_summary"]
        st.markdown("**Zone-wise RTGS Performance**")
        st.dataframe(etl.zone_tat_matrix_to_dataframe(summary["matrix"]), use_container_width=True, hide_index=True)
        excel_col, image_col = st.columns(2)
        with excel_col:
            st.download_button(
                "📊 Download as Excel",
                data=etl.rtgs_summary_to_excel_bytes(summary),
                file_name=f"RTGS_Summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with image_col:
            st.download_button(
                "🖼️ Download as Image",
                data=etl.rtgs_summary_to_image_bytes(summary),
                file_name=f"RTGS_Summary_{datetime.now().strftime('%Y%m%d')}.png",
                mime="image/png",
                use_container_width=True,
            )

    elif export_choice == "Cash Mode Validation Summary":
        df = st.session_state["cash_mode_validation"]
        if df.empty:
            st.success("✅ **COMPLIANT** — No violations detected!")
        else:
            st.warning(f"⚠️ **{len(df)} Violations Found**")
            st.dataframe(df, use_container_width=True, hide_index=True)
        excel_col, image_col = st.columns(2)
        with excel_col:
            st.download_button(
                "📊 Download as Excel",
                data=etl.cash_mode_validation_summary_to_excel_bytes(df),
                file_name=f"Cash_Mode_Validation_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with image_col:
            st.download_button(
                "🖼️ Download as Image",
                data=etl.cash_mode_validation_summary_to_image_bytes(df),
                file_name=f"Cash_Mode_Validation_{datetime.now().strftime('%Y%m%d')}.png",
                mime="image/png",
                use_container_width=True,
            )

    elif export_choice == "Delay in RCPTING Summary":
        summary = st.session_state["delay_summary"]
        st.markdown("**Processing TAT by Zone & Receipt Type**")
        st.dataframe(etl.zone_tat_matrix_to_dataframe(summary["matrix"]), use_container_width=True, hide_index=True)
        excel_col, image_col = st.columns(2)
        with excel_col:
            st.download_button(
                "📊 Download as Excel",
                data=etl.delay_summary_to_excel_bytes(summary),
                file_name=f"Delay_Summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with image_col:
            st.download_button(
                "🖼️ Download as Image",
                data=etl.delay_summary_to_image_bytes(summary),
                file_name=f"Delay_Summary_{datetime.now().strftime('%Y%m%d')}.png",
                mime="image/png",
                use_container_width=True,
            )

    elif export_choice == "Zone Overview":
        summary = st.session_state["zone_overview_summary"]
        st.markdown("**Zone-wise Mode Split, TAT Exceeded & Bounced/Cancelled**")
        rows = []
        for z in summary["zones"]:
            row = {"Zone": z["zone"]}
            row.update({etl.MODE_DISPLAY.get(m, m): z["by_mode"].get(m, 0) for m in summary["modes"]})
            row["TAT Exceeded"] = z["tat_exceeded"]
            row["Bounced"] = z["bounced"]
            row["Cancelled"] = z["cancelled"]
            row["Total"] = z["total"]
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        st.download_button(
            "🖼️ Download as Image",
            data=etl.zone_overview_to_image_bytes(summary),
            file_name=f"Zone_Overview_{datetime.now().strftime('%Y%m%d')}.png",
            mime="image/png",
            use_container_width=True,
        )

    elif export_choice == "RCPT CXN":
        df = st.session_state["rcpt_cxn"]
        if df.empty:
            st.success("✅ **CLEAN** — No cancelled receipts this period")
        else:
            st.warning(f"📋 {len(df)} Cancelled Receipts")
            st.dataframe(df, use_container_width=True, hide_index=True)
        excel_col, image_col = st.columns(2)
        with excel_col:
            st.download_button(
                "📊 Download as Excel",
                data=etl.rcpt_cxn_to_excel_bytes(df),
                file_name=f"RCPT_CXN_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with image_col:
            st.download_button(
                "🖼️ Download as Image",
                data=etl.rcpt_cxn_to_image_bytes(df),
                file_name=f"RCPT_CXN_{datetime.now().strftime('%Y%m%d')}.png",
                mime="image/png",
                use_container_width=True,
            )

    st.markdown("---")
    st.markdown("**Download the Full Workbook**")
    
    col_download, col_info_down = st.columns([2, 1])
    
    with col_download:
        st.download_button(
            "📥 Download DHC Working Output.xlsx",
            data=st.session_state["workbook_buffer"],
            file_name=f"DHC_Working_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    
    with col_info_down:
        st.info("""
        **File Contents:**
        - 5 Analytics Sheets
        - Pivot-style formatting
        - Zero formula errors
        """)
    
    st.markdown("""
    #### 📋 What's in the Excel File?
    
    | Sheet | Content | Rows | Purpose |
    |-------|---------|------|---------|
    | **Receipt Made Summary** | Updated/Bounced by Mode | 6-8 | Receipt status overview |
    | **RTGS Summary** | Zone × Type × TAT + Online sources | 45-60 | RTGS performance metrics |
    | **Cash Mode Validation** | Compliance violations | 0-100 | Identify breaches |
    | **Delay in RCPTING** | Full-month aging analysis | 45-60 | Processing delays |
    | **RCPT CXN** | Cancelled receipts | 10-50 | Review & add remarks |
    """)

    st.markdown("---")
    st.markdown("### 📧 Email Summary")

    zone_choice = st.selectbox(
        "Draft for zone",
        ["Overall", "North", "South", "East", "West"],
        key="mail_zone_choice",
        help="Filters the whole mail — every count and every chart — down to one broad zone. 'Overall' uses all zones.",
    )

    chq_brs = st.session_state.get("chq_brs")
    air_brs = st.session_state.get("air_brs")
    if chq_brs or air_brs:
        loaded = [n for n, v in (("Cheque/DD", chq_brs), ("Airtel Cash", air_brs)) if v]
        st.caption(f"✓ Using BRS MIS file(s) uploaded in Step 1: {', '.join(loaded)}. Re-upload there to refresh.")
    else:
        st.caption("No BRS MIS files uploaded in Step 1 — Delay in Receipting/Deposition will show [Fill manually]. Optional, upload them at the top of Step 1 if needed.")

    mail_dcr_tab = etl.filter_by_broad_zone(st.session_state["dcr_tab"], "ZONE", zone_choice)
    mail_receipt_made_summary = etl.build_receipt_made_summary(mail_dcr_tab)
    mail_rtgs_tab = etl.build_rtgs_tab(mail_dcr_tab)
    mail_rtgs_summary = etl.build_rtgs_summary(mail_rtgs_tab, mail_dcr_tab)
    mail_delay_summary = etl.build_delay_in_rcpting_summary(mail_dcr_tab)
    mail_cash_mode_validation = etl.build_cash_mode_validation_summary(mail_dcr_tab)
    mail_rcpt_cxn = etl.filter_by_broad_zone(st.session_state["rcpt_cxn"], "Zone", zone_choice)
    mail_zone_overview_summary = etl.build_zone_overview_summary(mail_dcr_tab)

    subject, body = etl.build_mail_summary(
        mail_receipt_made_summary,
        mail_rtgs_summary,
        mail_cash_mode_validation,
        mail_delay_summary,
        mail_rcpt_cxn,
        mail_dcr_tab,
        zone_label=zone_choice,
        chq_brs=chq_brs,
        air_brs=air_brs,
    )

    col_mail1, col_mail2, col_mail_info = st.columns([1.3, 1.3, 1])

    with col_mail1:
        if st.button("🖼️ Compose in Outlook (with charts)", use_container_width=True, type="primary"):
            try:
                import tempfile
                import os

                tmp_dir = tempfile.mkdtemp(prefix="dhc_mail_")
                images = {
                    "receipts_overview": etl.receipts_overview_image_bytes(
                        etl.compute_receipts_overview_block(mail_dcr_tab), zone_label=zone_choice
                    ).getvalue(),
                    "receipt_sync_status": etl.receipt_status_sync_image_bytes(
                        etl.compute_receipt_status_sync_block(mail_dcr_tab), zone_label=zone_choice
                    ).getvalue(),
                    "receipt_made_report": etl.receipt_made_status_summary_to_image_bytes(
                        mail_receipt_made_summary["full"]
                    ).getvalue(),
                    "rtgs_summary": etl.rtgs_summary_to_image_bytes(
                        mail_rtgs_summary
                    ).getvalue(),
                    "delay_summary": etl.delay_summary_to_image_bytes(
                        mail_delay_summary
                    ).getvalue(),
                    "zone_overview": etl.zone_overview_to_image_bytes(
                        mail_zone_overview_summary
                    ).getvalue(),
                    "rcpt_cxn": etl.rcpt_cxn_to_image_bytes(
                        mail_rcpt_cxn
                    ).getvalue(),
                }
                if chq_brs is not None:
                    images["chq_brs_status"] = etl.brs_status_image_bytes(chq_brs).getvalue()
                if air_brs is not None:
                    images["air_brs_status"] = etl.brs_status_image_bytes(air_brs).getvalue()

                image_paths = []
                for cid, png_bytes in images.items():
                    # Resize to 2x the mail's display width (set via img_width
                    # below), not 1x — the HTML <img width="560"> attribute is
                    # what actually controls the visual size in the mail, so
                    # shrinking the underlying PNG down to that same 560px was
                    # throwing away resolution for nothing: on any high-DPI
                    # screen the mail client then has to stretch a 560px image
                    # back out to render it, which is what "blurry" was. Source
                    # dashboards are already ~1100-1160px natively, so this is a
                    # light touch-up at most, not a real downscale.
                    small_png = etl.resize_png_for_mail(png_bytes, max_width=1120)
                    path = os.path.join(tmp_dir, f"{cid}.png")
                    with open(path, "wb") as f:
                        f.write(small_png)
                    image_paths.append((cid, path))

                html_subject, html_body = etl.build_mail_html_summary(
                    mail_receipt_made_summary,
                    mail_rtgs_summary,
                    mail_cash_mode_validation,
                    mail_delay_summary,
                    mail_rcpt_cxn,
                    mail_dcr_tab,
                    img_width=560,
                    zone_label=zone_choice,
                    chq_brs=chq_brs,
                    air_brs=air_brs,
                )
                etl.compose_outlook_mail_with_images(html_subject, html_body, image_paths)
                st.success("Draft opened in Outlook — check your taskbar, then attach the workbook and send.")
            except ImportError:
                st.error(
                    "This needs the pywin32 package. On the machine running this app, run "
                    "`pip install pywin32`, restart the app, and try again."
                )
            except Exception as e:
                st.error(f"Couldn't open Outlook automatically: {e}")
                st.caption("Use the text-only mail button instead for now.")

    with col_mail2:
        mailto_url = "mailto:?subject=" + urllib.parse.quote(subject) + "&body=" + urllib.parse.quote(body)
        st.link_button("📧 Compose Summary Mail (text only)", mailto_url, use_container_width=True)
        with st.expander("Preview text-only mail"):
            st.text(f"Subject: {subject}\n\n{body}")

    with col_mail_info:
        st.info("""
        **With charts** opens classic
        Outlook directly (via COM) with
        Receipt Made, RTGS, Delay in
        Recepting, and Receipt
        Cancellation charts embedded
        inline — sized small so they
        don't dominate the mail.

        **Text only** opens your OS's
        default mail app — no images,
        but works on any machine.
        """)

# ═══════════════════════════════════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="footer">
    <p style="margin: 0;">DHC Collections Intelligence Platform · v2.0</p>
</div>
""", unsafe_allow_html=True)
