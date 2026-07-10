import urllib.parse
import streamlit as st
import etl as etl
from datetime import datetime

st.set_page_config(page_title="Collection Operations Dashboard", page_icon="💰", layout="wide")

st.title("💰 Collection Operations Dashboard")
st.caption("Select a feature, upload the file, and click process.")

feature_selection = st.selectbox(
    "Select Feature to Process:",
    ["Pay Mode Exceeding", "Airtel Gateway Collection", "BRS Status"]
)

if feature_selection == "BRS Status":
    uploaded = st.file_uploader("Upload BRS MIS file (.xlsx)", type=["xlsx"])
else:
    uploaded = st.file_uploader("Upload DCR (.xlsb)", type=["xlsb"])

if not uploaded:
    st.info("Upload a file to begin.")
    st.session_state.pop("tat_result", None)
    st.stop()

# =========================================================
# FEATURE 1: PAY MODE EXCEEDING (TAT Validation)
# =========================================================
if feature_selection == "Pay Mode Exceeding":
    st.subheader("TAT Validation - Pay Mode Exceeding")
    mode = st.radio("Validation filter", ["Cash", "Cheque", "DD"], horizontal=True)

    current_signature = (uploaded.name, uploaded.size, mode, feature_selection)
    process_clicked = st.button("▶️ Click to Process", type="primary", use_container_width=True)

    if process_clicked:
        try:
            df = etl.load_dcr_raw(uploaded)
        except Exception as exc:
            st.error(f"Couldn't read this file: {exc}")
            st.session_state.pop("tat_result", None)
            st.stop()

        try:
            result_df, meta = etl.filter_mode_exceeding(df, mode)
        except ValueError as exc:
            st.error(str(exc))
            st.session_state.pop("tat_result", None)
            st.stop()

        st.session_state["tat_result"] = {
            "signature": current_signature,
            "result_df": result_df,
            "meta": meta,
        }

    stored = st.session_state.get("tat_result")

    if stored is None:
        st.info("Upload done — click **Click to Process** above to run the TAT check.")
        st.stop()

    if stored["signature"] != current_signature:
        st.warning(
            "You've changed the file, feature, or payment mode since the last run. "
            "Click **Click to Process** again to refresh the results below."
        )
        st.stop()

    result_df = stored["result_df"]
    meta = stored["meta"]

    st.caption(
        f"TAT for **{mode}**: **{meta['threshold']} day(s)**. "
        f"PENDING DAYS computed using {meta['present_source']}. "
        f"{meta['total_matched']:,} {mode} receipt(s) in this file, {meta['total_exceeding']:,} exceeding TAT."
    )
    if meta["missing_columns"]:
        st.warning(
            "Couldn't find these expected columns in the file (shown as 'N/A' below) — "
            "check the exact header spelling in your DCR against etl.OUTPUT_COLUMNS: "
            + ", ".join(meta["missing_columns"])
        )
    if meta.get("unparseable_dates"):
        st.warning(
            f"{meta['unparseable_dates']:,} {mode} row(s) had a Date or RECEIPT ENTER DATE that "
            "couldn't be parsed and were excluded from the TAT check."
        )

    st.metric(f"{mode} receipts exceeding TAT", meta["total_exceeding"])

    if result_df.empty:
        st.success(f"No {mode} receipts exceeding TAT right now — nothing to send.")
        st.stop()

    st.dataframe(result_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 📧 Email This List")

    subject_html, html_body = etl.build_tat_mail_html(mode, result_df)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🖼️ Compose in Outlook (full list)", use_container_width=True):
            try:
                etl.compose_outlook_mail_html(subject_html, html_body)
                st.success("Draft opened in Outlook — check your taskbar, review, then send.")
            except ImportError:
                st.error(
                    "This needs the pywin32 package. On the machine running this app, run "
                    "`pip install pywin32`, restart the app, and try again."
                )
            except Exception as exc:
                st.error(f"Couldn't open Outlook automatically: {exc}")

    with col2:
        subject_text, body_text = etl.build_tat_mail_text(mode, result_df)
        mailto_url = "mailto:?subject=" + urllib.parse.quote(subject_text) + "&body=" + urllib.parse.quote(body_text)
        st.link_button("📧 Compose (text only, first 20 rows)", mailto_url, use_container_width=True)
        with st.expander("Preview text-only mail"):
            st.text(f"Subject: {subject_text}\n\n{body_text}")


# =========================================================
# FEATURE 2: AIRTEL GATEWAY COLLECTION
# =========================================================
# =========================================================
# FEATURE 2: AIRTEL GATEWAY COLLECTION
# =========================================================
# =========================================================
# FEATURE 2: AIRTEL GATEWAY COLLECTION
# =========================================================
# =========================================================
# FEATURE 2: AIRTEL GATEWAY COLLECTION
# =========================================================
elif feature_selection == "Airtel Gateway Collection":
    st.subheader("Airtel Money Gateway Collection Dashboard")
    
    process_clicked = st.button("▶️ Generate Dashboard", type="primary", use_container_width=True)
    
    if process_clicked or "airtel_zone_df" in st.session_state:
        if process_clicked:
            try:
                df = etl.load_dcr_raw(uploaded)
                zone_df, region_df = etl.generate_airtel_dashboard(df)
                
                st.session_state["airtel_zone_df"] = zone_df
                st.session_state["airtel_region_df"] = region_df
            except Exception as exc:
                st.error(f"Error processing the Airtel Dashboard: {exc}")
                st.stop()
        
        zone_df = st.session_state["airtel_zone_df"]
        region_df = st.session_state["airtel_region_df"]

        # Calculate the Bank Percentage dynamically from the Grand Total row
        grand_total_bank = zone_df.loc[zone_df[('Zone', '')] == 'Grand Total', ('Deposited - Bank', 'Value')].values[0]
        grand_total_overall = zone_df.loc[zone_df[('Zone', '')] == 'Grand Total', ('Overall', 'Value')].values[0]
        bank_pct = (grand_total_bank / grand_total_overall * 100) if grand_total_overall > 0 else 0

        st.warning(f"⚠️ **{bank_pct:.2f}%** of the Cash got Deposited in Bank.")

        def style_df(df):
            format_dict = {}
            for col in df.columns:
                if isinstance(col, tuple):
                    if col[1] in ['Count %', 'Value %']:
                        format_dict[col] = "{:,.2f}%"
                    elif col[1] == 'Count':
                        format_dict[col] = "{:,.0f}"
                    elif col[1] == 'Value':
                        format_dict[col] = "{:,.2f}"
            return df.style.format(format_dict)

        # Swapped Order: Zone First, Region Second
        st.write("### Zone-wise Dashboard")
        st.dataframe(style_df(zone_df), use_container_width=True, hide_index=True)
        
        st.divider()
        
        st.write("### Region-wise Dashboard")
        st.dataframe(style_df(region_df), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### 📧 Email This Dashboard")

        if st.button("🖼️ Draft Airtel Dashboard in Outlook", use_container_width=True):
            try:
                subject_html, html_body = etl.build_airtel_mail_html(zone_df, region_df, bank_pct)
                etl.compose_outlook_mail_html(subject_html, html_body)
                st.success("Draft opened in Outlook — check your taskbar, review, then send.")
            except ImportError:
                st.error("Missing pywin32 package. Run `pip install pywin32`.")
            except Exception as exc:
                st.error(f"Couldn't open Outlook automatically: {exc}")

# =========================================================
# FEATURE 3: BRS STATUS (Cheque/DD and Airtel)
# =========================================================
elif feature_selection == "BRS Status":
    st.subheader("BRS Status Dashboard")

    brs_mode = st.radio(
        "Select BRS Type",
        ["Cheque", "Airtel"],
        horizontal=True,
        help="Cheque = CHEQUE/DD BRS Status | Airtel = Airtel Cash BRS Status",
    )

    report_date = st.text_input(
        "Report Date (for title and mail subject)",
        value=datetime.now().strftime("%d-%m-%Y"),
    )

    current_sig = (uploaded.name, uploaded.size, brs_mode, feature_selection)
    process_clicked = st.button("▶️ Generate BRS Status", type="primary", use_container_width=True)

    if process_clicked:
        try:
            df = etl.load_brs_raw(uploaded)
        except Exception as exc:
            st.error(f"Couldn't read this file: {exc}")
            st.stop()
        try:
            result = etl.process_brs(df, brs_mode, report_date)
            st.session_state["brs_result"] = {"sig": current_sig, "result": result}
        except Exception as exc:
            st.error(f"Error processing BRS: {exc}")
            st.stop()

    stored = st.session_state.get("brs_result")
    if stored is None:
        st.info("Upload a BRS MIS Excel file, then click **Generate BRS Status**.")
        st.stop()
    if stored["sig"] != current_sig:
        st.warning("File or settings changed — click **Generate BRS Status** to refresh.")
        st.stop()

    result = stored["result"]
    st.success(f"Processed: {result['title']} — {sum(result['grand'].values())} total receipts")

    # --- Download Excel ---
    st.download_button(
        label="📥 Download BRS Status Excel",
        data=result["excel_bytes"],
        file_name=f"BRS_Status_{brs_mode}_{report_date.replace('/','-')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # --- Chart preview ---
    st.markdown("### 📊 Visual Summary")
    st.image(result["chart_bytes"], use_container_width=True)

    # --- Email ---
    st.markdown("---")
    st.markdown("### 📧 Email This Summary")

    if st.button("🖼️ Compose in Outlook (with chart)", use_container_width=True, type="primary"):
        try:
            import tempfile, os
            subject, html_body = etl.build_brs_mail_html(
                brs_mode, result["pivot"], result["grand"],
                result["buckets"], result["ops_rows"], report_date,
                chart_cid="brs_chart",
            )
            tmp = tempfile.mkdtemp(prefix="brs_mail_")
            chart_path = os.path.join(tmp, "brs_chart.png")
            with open(chart_path, "wb") as f:
                f.write(result["chart_bytes"])
            etl.compose_outlook_mail_with_images(subject, html_body, [("brs_chart", chart_path)])
            st.success("Draft opened in Outlook — attach the Excel, review, then send.")
        except ImportError:
            st.error("Needs pywin32: run `pip install pywin32` and restart.")
        except Exception as exc:
            st.error(f"Couldn't open Outlook: {exc}")