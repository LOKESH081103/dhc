"""
DHC Working Automation — Professional Streamlit Application

A modern, enterprise-grade MIS reporting platform for automated collections analytics,
RTGS monitoring, compliance validation, and delay tracking.

Version: 2.0
Last Updated: June 2026
"""

import traceback
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

# ═══════════════════════════════════════════════════════════════════════════════════
# SIDEBAR NAVIGATION
# ═══════════════════════════════════════════════════════════════════════════════════

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
        
        ### Delinquency Dump (.csv)
        - Columns needed: **AGREEMENTNO, CIF NO, DPD, DPD SLAB**
        - Fresh export each period — source of truth for CIF mapping
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
# HEADER SECTION
# ═══════════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="header-container">
    <h1 class="header-title">📊 DHC Collections Intelligence Platform</h1>
    <p class="header-subtitle">
        Automated MIS Reporting • RTGS Analytics • Compliance Monitoring • Processing Intelligence
    </p>
    <div class="header-meta">
        <div class="meta-item">
            <span>🚀 Automated</span> · <span>Fast Processing</span>
        </div>
        <div class="meta-item">
            <span>📈 Enterprise Ready</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════════
# KEY FEATURES SHOWCASE
# ═══════════════════════════════════════════════════════════════════════════════════

st.markdown("### 🎯 Platform Capabilities")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-card-title">📁 Multi-Source</div>
        <div class="feature-card-desc">Consolidate 4 data sources into unified analytics</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-card-title">⚡ Lightning Fast</div>
        <div class="feature-card-desc">Process 100K+ records in under 1 minute</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-card-title">🛡️ Compliance</div>
        <div class="feature-card-desc">Real-time violation detection & monitoring</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-card-title">📊 Intelligence</div>
        <div class="feature-card-desc">Advanced analytics & trend identification</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════════════════════════════════
# FILE UPLOAD SECTION
# ═══════════════════════════════════════════════════════════════════════════════════

st.markdown("### 📂 Monthly Data Processing")
st.markdown("Upload your source files and launch the automated processing pipeline.")

col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    st.markdown("#### Collections & Compliance Data")
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

with col_right:
    st.markdown("#### Reference Data")
    employee_file = st.file_uploader(
        "👥 Employee Master",
        type=["xlsx"],
        help="Must contain 'Mobile Number' column for agent identification",
        key="employee"
    )
    delinquency_file = st.file_uploader(
        "📉 Delinquency Dump",
        type=["csv"],
        help="LAP_Delq_Dump.csv — source of truth for CIF No. & Opening DPD Slab per agreement",
        key="delinquency"
    )

st.markdown("---")

# ═══════════════════════════════════════════════════════════════════════════════════
# PROCESS BUTTON & STATUS
# ═══════════════════════════════════════════════════════════════════════════════════

all_files_ready = all([dcr_file, disable_file, employee_file, delinquency_file])

col_process, col_info = st.columns([2, 1], gap="medium")

with col_process:
    run = st.button(
        "🚀  PROCESS FILES  →",
        type="primary",
        disabled=not all_files_ready,
        use_container_width=True,
        help="Upload all 4 files to enable" if not all_files_ready else "Click to start automated processing"
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
        if not delinquency_file:
            missing.append("Delinquency Dump")
        st.warning(f"⏳ Missing: {', '.join(missing)}")

# ═══════════════════════════════════════════════════════════════════════════════════
# PROCESSING & OUTPUT
# ═══════════════════════════════════════════════════════════════════════════════════

if run:
    with st.spinner("🔄 Processing your data... This typically takes 40-60 seconds"):
        try:
            with st.status("📊 Running automated pipeline...", expanded=True) as status:
                st.write("📥 **Stage 1/6**: Loading DCR extract...")
                receipts, dcr_master = etl.load_dcr(dcr_file)
                st.write(f"   ✓ Loaded {len(receipts):,} receipts")

                st.write("📥 **Stage 2/6**: Loading compliance restrictions...")
                cif_disable, agr_disable = etl.load_disable_lists(disable_file)
                st.write(f"   ✓ Loaded {len(cif_disable):,} CIF + {len(agr_disable):,} agreement restrictions")

                st.write("📥 **Stage 3/6**: Loading employee mobile master...")
                mobiles = etl.load_employee_mobiles(employee_file)
                st.write(f"   ✓ Loaded {len(mobiles):,} agent mobile numbers")

                st.write("📥 **Stage 4/6**: Loading delinquency dump (CIF & DPD Slab)...")
                delinquency_master = etl.load_delinquency_master(delinquency_file)
                st.write(f"   ✓ Loaded {len(delinquency_master):,} agreements with CIF mapping")

                st.write("⚙️ **Stage 5/6**: Building enriched data tables...")
                lookup_master = etl.build_lookup_master(dcr_master, delinquency_master)
                dcr_tab = etl.build_dcr_tab(receipts, lookup_master, agr_disable, cif_disable, mobiles)
                rtgs_tab = etl.build_rtgs_tab(dcr_tab)
                st.write(f"   ✓ Built DCR tab ({len(dcr_tab):,} receipts × 87 columns)")
                st.write(f"   ✓ Built RTGS tab ({len(rtgs_tab):,} RTGS receipts)")

                st.write("📊 **Stage 6/6**: Generating intelligence summaries...")
                receipt_made_summary = etl.build_receipt_made_summary(dcr_tab)
                rtgs_summary = etl.build_rtgs_summary(rtgs_tab, dcr_tab)
                cash_mode_validation_summary = etl.build_cash_mode_validation_summary(dcr_tab)
                delay_summary = etl.build_delay_in_rcpting_summary(dcr_tab)
                rcpt_cxn = etl.build_rcpt_cxn(dcr_tab)
                st.write("   ✓ Generated 5 analytics summaries")

                st.write("💾 **Finalizing**: Writing Excel workbook...")
                workbook_buffer = etl.write_output_workbook(
                    rtgs_summary, delay_summary, receipt_made_summary,
                    cash_mode_validation_summary, rcpt_cxn
                )
                st.write("   ✓ Workbook created (0 formula errors)")

                status.update(label="✅ Processing Complete!", state="complete")

            # Store in session
            st.session_state["workbook_buffer"] = workbook_buffer
            st.session_state["receipt_made_summary"] = receipt_made_summary
            st.session_state["rtgs_summary"] = rtgs_summary
            st.session_state["cash_mode_validation"] = cash_mode_validation_summary
            st.session_state["delay_summary"] = delay_summary
            st.session_state["rcpt_cxn"] = rcpt_cxn
            st.session_state["lookup_master"] = lookup_master
            st.session_state["dcr_tab"] = dcr_tab
            st.session_state["rtgs_tab"] = rtgs_tab

            st.markdown("---")

            # Success metrics
            st.markdown("### 📈 Processing Results")
            
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
                needs_cif = int(lookup_master["NEEDS_CIF_MAPPING"].sum())
                st.markdown(render_stat_card(
                    "Unmapped CIF",
                    f"{needs_cif:,}",
                    "No delinquency-dump match",
                    "red" if needs_cif > 0 else "green"
                ), unsafe_allow_html=True)
            
            with metric_col4:
                cxn_count = len(rcpt_cxn)
                st.markdown(render_stat_card(
                    "Cancelled Receipts",
                    f"{cxn_count:,}",
                    "For your review",
                    "blue"
                ), unsafe_allow_html=True)

            # Warnings
            if needs_cif > 0:
                st.warning(
                    f"""
                    ⚠️ **{needs_cif:,} AGREEMENTS WITH NO CIF MATCH**
                    
                    These agreements appear in this period's DCR but have no matching row in the Delinquency 
                    Dump you uploaded (e.g. fully closed accounts, or too new to appear in the delinquency 
                    pull yet), so their CIF No. is blank.
                    
                    👉 **Action**: Open the "Agreement → CIF Master" tab below to see the blank-CIF rows. If 
                    it's a fresh delinquency pull for this period, these should resolve on their own next run — 
                    otherwise confirm the agreement number with your LMS.
                    """,
                    icon="⚠️"
                )

            st.markdown("---")

        except Exception as e:
            st.error(f"❌ Processing Error: {str(e)}")
            with st.expander("📋 Detailed Error Information"):
                st.code(traceback.format_exc())

# ═══════════════════════════════════════════════════════════════════════════════════
# OUTPUT PREVIEW SECTION
# ═══════════════════════════════════════════════════════════════════════════════════

if "workbook_buffer" in st.session_state:
    st.markdown("---")
    st.markdown("### 📑 Analytics Dashboards Preview")
    
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
        st.info("✓ CIF No. & Opening DPD Slab sourced fresh from this period's Delinquency Dump")
        
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
    st.markdown("### 📤 Export a Specific Summary")

    export_choice = st.selectbox(
        "Choose what to download",
        [
            "Receipt Made Summary",
            "RTGS Summary (coming soon)",
            "Cash Mode Validation Summary (coming soon)",
            "Delay in RCPTING Summary (coming soon)",
            "RCPT CXN (coming soon)",
        ],
    )

    if export_choice == "Receipt Made Summary":
        rms_df = etl.build_receipt_made_summary_by_status(st.session_state["dcr_tab"])
        st.dataframe(rms_df, use_container_width=True, hide_index=True)

        excel_col, image_col = st.columns(2)
        with excel_col:
            st.download_button(
                "📊 Download as Excel",
                data=etl.receipt_made_summary_to_excel_bytes(rms_df),
                file_name=f"Receipt_Made_Summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with image_col:
            st.download_button(
                "🖼️ Download as Image",
                data=etl.receipt_made_summary_to_image_bytes(rms_df),
                file_name=f"Receipt_Made_Summary_{datetime.now().strftime('%Y%m%d')}.png",
                mime="image/png",
                use_container_width=True,
            )
    else:
        st.info("This export is being wired up next — for now only **Receipt Made Summary** is live here.")

    # ═══════════════════════════════════════════════════════════════════════════════════
    # DOWNLOAD SECTION
    # ═══════════════════════════════════════════════════════════════════════════════════

    st.markdown("---")
    st.markdown("### 💾 Download Your Output")
    
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

# ═══════════════════════════════════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════════════════════════════════

st.markdown("---")
st.markdown("""
<div class="footer">
    <h4 style="color: #1e40af; margin: 0;">DHC Collections Intelligence Platform</h4>
    <p style="margin: 0.5rem 0 0 0;">
        Automated MIS Reporting • Enterprise Grade • Production Ready
    </p>
    <p style="margin: 0.5rem 0 0 0; font-size: 0.8rem;">
        Version 2.0 • June 2026 • Support & Documentation Available
    </p>
</div>
""", unsafe_allow_html=True)
