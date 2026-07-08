import urllib.parse

import streamlit as st

import etl

st.set_page_config(page_title="Cash / Cheque / DD — TAT Validation", page_icon="\U0001F4B0", layout="wide")

st.title("\U0001F4B0 Cash / Cheque / DD \u2014 TAT Validation")
st.caption(
    "Upload the DCR, pick a payment mode, then hit Process. See every receipt that's "
    "blown past its TAT (Cash: 1 day \u00b7 Cheque/DD: 3 days), then email the list straight from here."
)

uploaded = st.file_uploader("Upload DCR (.xlsb)", type=["xlsb"])
mode = st.radio("Validation filter", ["Cash", "Cheque", "DD"], horizontal=True)

if not uploaded:
    st.info("Upload a DCR file to begin.")
    st.session_state.pop("tat_result", None)
    st.stop()

# Signature of "what would be processed right now" — used to tell whether
# the stored result (if any) still matches the current file + mode, or is
# stale because the user picked a new file / switched modes since Process
# was last clicked.
current_signature = (uploaded.name, uploaded.size, mode)

process_clicked = st.button("\u25B6\uFE0F Click to Process", type="primary", use_container_width=True)

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
    st.info("Upload done \u2014 click **Click to Process** above to run the TAT check.")
    st.stop()

if stored["signature"] != current_signature:
    st.warning(
        "You've changed the file or the payment mode since the last run. "
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
        "Couldn't find these expected columns in the file (shown as 'N/A' below) \u2014 "
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
    st.success(f"No {mode} receipts exceeding TAT right now \u2014 nothing to send.")
    st.stop()

st.dataframe(result_df, use_container_width=True, hide_index=True)

st.markdown("---")
st.markdown("### \U0001F4E7 Email This List")

subject_html, html_body = etl.build_tat_mail_html(mode, result_df)

col1, col2 = st.columns(2)
with col1:
    if st.button("\U0001F5BC\uFE0F Compose in Outlook (full list)", use_container_width=True):
        try:
            etl.compose_outlook_mail_html(subject_html, html_body)
            st.success("Draft opened in Outlook \u2014 check your taskbar, review, then send.")
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
    st.link_button("\U0001F4E7 Compose (text only, first 20 rows)", mailto_url, use_container_width=True)
    with st.expander("Preview text-only mail"):
        st.text(f"Subject: {subject_text}\n\n{body_text}")