"""
Agreement Address Quality Checker - 4-layer pipeline
-----------------------------------------------------
Layer 1: Structural/mechanical rules (free, offline, instant)
Layer 2: Pincode master-data validation (free public API, needs internet)
Layer 3: Placeholder / gibberish / foreign-location dictionary (free, offline)
Layer 4: Optional AI semantic judge - Gemini free tier (needs your API key)

Run:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import re
import time

import pandas as pd
import streamlit as st

from pincode_lookup import lookup_pincode
from ai_review import gemini_check
from feedback_store import load_feedback, save_feedback, normalize, previously_cleared_addresses

st.set_page_config(page_title="Agreement Address Quality Checker", layout="wide")

# ----------------------------------------------------------------------
# Reference data
# ----------------------------------------------------------------------
INDIAN_STATES = [
    "ANDHRA PRADESH", "ARUNACHAL PRADESH", "ASSAM", "BIHAR", "CHHATTISGARH",
    "GOA", "GUJARAT", "HARYANA", "HIMACHAL PRADESH", "JHARKHAND", "KARNATAKA",
    "KERALA", "MADHYA PRADESH", "MAHARASHTRA", "MANIPUR", "MEGHALAYA",
    "MIZORAM", "NAGALAND", "ODISHA", "PUNJAB", "RAJASTHAN", "SIKKIM",
    "TAMIL NADU", "TELANGANA", "TRIPURA", "UTTAR PRADESH", "UTTARAKHAND",
    "WEST BENGAL", "DELHI", "JAMMU AND KASHMIR", "LADAKH", "PUDUCHERRY",
    "CHANDIGARH", "ANDAMAN AND NICOBAR", "DADRA AND NAGAR HAVELI",
    "DAMAN AND DIU", "LAKSHADWEEP",
]

COMMON_SAFE_LONG_WORDS = {"MAHARASHTRA", "TELANGANA", "CHHATTISGARH", "PONDICHERRY", "VISAKHAPATNAM"}

PLACEHOLDER_PHRASES = {
    "NA", "N A", "N/A", "N.A", "N.A.", "NIL", "NONE", "XXX", "XXXX", "XYZ", "ABC",
    "TEST", "TESTING", "TBD", "PENDING", "DUMMY", "SAMPLE", "DEFAULT", "UNKNOWN",
    "NOT AVAILABLE", "ADDRESS NOT AVAILABLE", "SAME AS ABOVE", "SAME AS PREVIOUS",
    "ASDF", "ASDFGH", "QWERTY",
}
PLACEHOLDER_WORDS = {"TEST", "TESTING", "TBD", "DUMMY", "SAMPLE", "ASDF", "ASDFGH", "QWERTY", "XYZ", "NIL"}

FOREIGN_LOCATION_HINTS = {
    "DUBAI", "UAE", "ABU DHABI", "SHARJAH", "SINGAPORE", "LONDON", "USA",
    "UNITED STATES", "UNITED KINGDOM", "CANADA", "AUSTRALIA", "NEPAL", "DOHA", "QATAR",
}

CRITICAL_ISSUE_PREFIXES = {
    "EMPTY_ADDRESS", "MISSING_PINCODE", "PINCODE_NOT_FOUND_IN_INDIA",
    "PLACEHOLDER_ADDRESS", "ADDRESS_TOO_SHORT",
}

ISSUE_DESCRIPTIONS = {
    "EMPTY_ADDRESS": "Address field is blank",
    "DOUBLE_COMMA_EMPTY_FIELD": "Contains ',,' - an empty field between commas",
    "PINCODE_DUPLICATED": "6-digit pincode appears twice back-to-back",
    "PINCODE_GLUED_TO_TEXT": "Pincode is stuck directly to a word with no space",
    "MISSING_PINCODE": "No 6-digit pincode found",
    "STATE_NOT_FOUND": "No recognizable Indian state name in the address",
    "ADDRESS_TOO_SHORT": "Address has very few words - likely incomplete",
    "POSSIBLE_MERGED_WORDS": "A long word may be two+ words stuck together",
    "HOUSE_NO_ZERO_OR_PLACEHOLDER": "House/flat number looks like a placeholder",
    "PINCODE_NOT_FOUND_IN_INDIA": "Pincode doesn't exist in the official India Post database",
    "PINCODE_STATE_MISMATCH": "Pincode belongs to a different state than what's written",
    "PLACEHOLDER_ADDRESS": "Entire address is a placeholder value (NA, TEST, etc.)",
    "PLACEHOLDER_WORD": "Contains a placeholder/junk word",
    "FOREIGN_LOCATION_MENTIONED": "Mentions a location outside India",
    "REPEATED_CHARACTER_RUN": "Same character repeated 4+ times in a row (e.g. aaaa)",
    "POSSIBLE_GIBBERISH_TEXT": "Long run of consonants suggests random/gibberish text",
    "AI_FLAGGED": "AI reviewer flagged this address",
}

DEMO_DATA = [
    ("AGR001", "ABHISHEK BUNGALOW NO. ONEKALPATARU NAGAR ASHOKA MARG , 422011"),
    ("AGR002", "SECTOR NO-4,CBD BELAPUR , NAVI MUMBAI400206"),
    ("AGR003", "FLAT NO- X, 5 TH FLOOR, BEACON CHSSOUTH AVENUEOPP RAMKRISHNA MISSION HOSPITAL, , SANTACRUZ-W, MUMBAI- 400054400054"),
    ("AGR004", "# 0, INSIDE NEW MARKET BAGGA MARKET , ,JAGADHRI YAMUNA NAGAR HARYANA - 135001"),
    ("AGR005", "# INDUSTRIEL AREA, , NEAR JODI FNAST ROAD YAMUNA NAGAR HARYANA - 135002"),
    ("AGR006", "# CHHACHHROULI ROAD, JAGADHRI, , YAMUNA NAGAR HARYANA - 135002"),
    ("AGR007", "YELAMANCHILI ROADATCHUT,APURAM, MAIN ROAD , ,MAIN ROAD531011"),
    ("AGR008", "12, GREEN PARK EXTENSION, NEW DELHI, DELHI - 110016"),
    ("AGR009", "MAIN ROAD 1, DUBAI"),
    ("AGR010", "NA"),
    ("AGR011", "FLAT 302 SUNRISE APARTMENTS MG ROAD BANGALORE KARNATAKA - 999999"),
]


def describe_issue(issue: str) -> str:
    base = issue.split("(")[0]
    return ISSUE_DESCRIPTIONS.get(base, base)


def severity_for(issue_codes):
    if not issue_codes:
        return "Clean"
    if any(i.split("(")[0] in CRITICAL_ISSUE_PREFIXES for i in issue_codes):
        return "Critical"
    return "Warning"


# ----------------------------------------------------------------------
# Layer 1 - structural rules
# ----------------------------------------------------------------------
def layer1_structural(addr: str, tokens, min_words: int, merge_len_threshold: int):
    issues = []
    if re.search(r",\s*,", addr):
        issues.append("DOUBLE_COMMA_EMPTY_FIELD")
    if re.search(r"(\d{6})\1", addr):
        issues.append("PINCODE_DUPLICATED")
    glued_match = re.search(r"[A-Za-z](\d{6})\b", addr)
    if glued_match and "PINCODE_DUPLICATED" not in issues:
        issues.append("PINCODE_GLUED_TO_TEXT")

    if not any(state in addr.upper() for state in INDIAN_STATES):
        issues.append("STATE_NOT_FOUND")

    phrase_counts = {}
    for n in (2, 3):
        for i in range(len(tokens) - n + 1):
            phrase = " ".join(tokens[i:i + n])
            phrase_counts[phrase] = phrase_counts.get(phrase, 0) + 1
    repeated = [p for p, c in phrase_counts.items() if c > 1 and len(p) > 6]
    if repeated:
        issues.append(f"REPEATED_PHRASE({repeated[0]})")

    if len(tokens) < min_words:
        issues.append("ADDRESS_TOO_SHORT")

    long_tokens = [t for t in tokens if len(t) >= merge_len_threshold and t not in COMMON_SAFE_LONG_WORDS]
    if long_tokens:
        issues.append(f"POSSIBLE_MERGED_WORDS({long_tokens[0]})")

    if re.search(r"#\s*0\b", addr):
        issues.append("HOUSE_NO_ZERO_OR_PLACEHOLDER")

    return issues


# ----------------------------------------------------------------------
# Layer 2 - pincode master-data validation
# ----------------------------------------------------------------------
def extract_pins(addr: str):
    pins = set(re.findall(r"\b\d{6}\b", addr))
    glued_pins = set(re.findall(r"[A-Za-z](\d{6})\b", addr))
    return pins | glued_pins


def layer2_pincode_master(addr_upper: str, pins: set, enabled: bool):
    """Returns (issues, network_status). network_status in {skipped, ok, error}."""
    if not enabled or not pins:
        return [], "skipped"
    issues = []
    network_status = "ok"
    for pin in sorted(pins):
        result = lookup_pincode(pin)
        if result is None:
            issues.append("PINCODE_NOT_FOUND_IN_INDIA")
        elif result == "ERROR":
            network_status = "error"
        else:
            actual_state = str(result.get("state", "")).upper()
            if actual_state and actual_state not in addr_upper:
                other_states = [s for s in INDIAN_STATES if s in addr_upper and s != actual_state]
                if other_states:
                    issues.append(f"PINCODE_STATE_MISMATCH(pin={pin} actual={actual_state} stated={other_states[0]})")
    return issues, network_status


# ----------------------------------------------------------------------
# Layer 3 - placeholder / gibberish / foreign-location dictionary
# ----------------------------------------------------------------------
def layer3_placeholder_gibberish(addr_upper: str, tokens):
    issues = []
    stripped = re.sub(r"[^A-Z ]", " ", addr_upper)
    stripped = re.sub(r"\s+", " ", stripped).strip()
    if stripped in PLACEHOLDER_PHRASES:
        issues.append("PLACEHOLDER_ADDRESS")

    hit = set(tokens) & PLACEHOLDER_WORDS
    if hit and "PLACEHOLDER_ADDRESS" not in issues:
        issues.append(f"PLACEHOLDER_WORD({sorted(hit)[0]})")

    foreign_hit = [f for f in FOREIGN_LOCATION_HINTS if f in addr_upper]
    if foreign_hit:
        issues.append(f"FOREIGN_LOCATION_MENTIONED({foreign_hit[0]})")

    if re.search(r"([A-Za-z0-9])\1{3,}", addr_upper):
        issues.append("REPEATED_CHARACTER_RUN")

    if re.search(r"[BCDFGHJKLMNPQRSTVWXYZ]{6,}", addr_upper):
        issues.append("POSSIBLE_GIBBERISH_TEXT")

    return issues


# ----------------------------------------------------------------------
# Orchestration
# ----------------------------------------------------------------------
def analyze_address(addr, min_words, merge_len_threshold, layer2_enabled):
    if not isinstance(addr, str) or not addr.strip():
        return ["EMPTY_ADDRESS"], "skipped"

    addr = addr.strip()
    addr_upper = addr.upper()
    tokens = re.findall(r"[A-Za-z]+", addr_upper)
    pins = extract_pins(addr)

    issues = []
    issues += layer1_structural(addr, tokens, min_words, merge_len_threshold)
    issues += layer3_placeholder_gibberish(addr_upper, tokens)

    if not pins and "MISSING_PINCODE" not in issues:
        issues.append("MISSING_PINCODE")

    l2_issues, net_status = layer2_pincode_master(addr_upper, pins, layer2_enabled)
    issues += l2_issues

    seen = set()
    deduped = []
    for i in issues:
        if i not in seen:
            seen.add(i)
            deduped.append(i)

    return deduped, net_status


def process_dataframe(df, address_col, agreement_col, min_words, merge_len_threshold,
                       layer2_enabled, cleared_addresses):
    rows = []
    net_error_seen = False
    for _, row in df.iterrows():
        addr = row[address_col]
        issues, net_status = analyze_address(addr, min_words, merge_len_threshold, layer2_enabled)
        if net_status == "error":
            net_error_seen = True

        cleared = isinstance(addr, str) and normalize(addr) in cleared_addresses
        if cleared:
            issues = []

        rows.append({
            "Agreement No": row[agreement_col] if agreement_col else "",
            "Address": addr,
            "Severity": severity_for(issues) if not cleared else "Clean (reviewer-cleared)",
            "Issue Count": len(issues),
            "Issues": "; ".join(issues) if issues else "",
            "Issue Details": "; ".join(describe_issue(i) for i in issues) if issues else "",
        })
    return pd.DataFrame(rows), net_error_seen


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
st.title("📋 Agreement Address Quality Checker")
st.caption("4-layer pipeline: structural rules -> pincode master-data check -> "
           "placeholder/gibberish dictionary -> optional AI semantic review.")

if "feedback_df" not in st.session_state:
    st.session_state.feedback_df = load_feedback()

with st.sidebar:
    st.header("⚙️ Layer 1 settings")
    min_words = st.slider("Minimum words for a valid address", 2, 10, 5)
    merge_len_threshold = st.slider("Merged-word length threshold", 8, 20, 12)

    st.divider()
    st.header("🌐 Layer 2: Pincode master-data")
    layer2_enabled = st.checkbox("Enable pincode validation (needs internet)", value=True)
    st.caption("Checks every pincode against the official India Post directory "
               "(free public API). Catches fake pincodes and state mismatches, "
               "e.g. 'Dubai' with no valid Indian pincode.")

    st.divider()
    st.header("📖 Layer 3: Placeholder dictionary")
    st.caption("Always on, free, offline. Catches NA/TEST/XXX-style junk, "
               "foreign city/country mentions, repeated characters, gibberish runs.")

    st.divider()
    st.header("🤖 Layer 4: Optional AI review (Gemini)")
    use_ai = st.checkbox("Enable AI semantic review")
    api_key = ""
    ai_scope = "Flagged rows only"
    sample_size = 20
    if use_ai:
        api_key = st.text_input("Gemini API key", type="password",
                                 help="Free key: https://aistudio.google.com/app/apikey")
        ai_scope = st.radio("Run AI on", ["Flagged rows only", "Random sample", "All rows"], index=0)
        if ai_scope == "Random sample":
            sample_size = st.number_input("Sample size", 5, 200, 20)
        st.caption("Your address data is sent to Google's API for this step only. "
                   "Free tier has rate limits, so requests are paced automatically.")

st.subheader("1. Load data")
col_a, col_b = st.columns([2, 1])
with col_a:
    uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx", "xls"])
with col_b:
    st.write("")
    st.write("")
    use_demo = st.button("▶ Try demo data instead")

df = None
if uploaded is not None:
    df = pd.read_excel(uploaded)
elif use_demo:
    df = pd.DataFrame(DEMO_DATA, columns=["Agreement No", "Address"])
    st.session_state["demo_loaded"] = True
elif st.session_state.get("demo_loaded"):
    df = pd.DataFrame(DEMO_DATA, columns=["Agreement No", "Address"])

if df is not None:
    st.success(f"Loaded {len(df)} rows.")
    st.dataframe(df.head(), use_container_width=True)

    st.subheader("2. Select columns")
    cols = list(df.columns)

    def guess(col_list, keyword):
        for c in col_list:
            if keyword in str(c).lower():
                return c
        return col_list[0]

    c1, c2 = st.columns(2)
    with c1:
        agreement_col = st.selectbox("Agreement No. column", cols, index=cols.index(guess(cols, "agree")))
    with c2:
        address_col = st.selectbox("Address column", cols, index=cols.index(guess(cols, "address")))

    if st.button("🔍 Run address check", type="primary"):
        cleared_addresses = previously_cleared_addresses(st.session_state.feedback_df)

        with st.spinner("Running Layer 1-3 checks..."):
            result_df, net_error_seen = process_dataframe(
                df, address_col, agreement_col, min_words, merge_len_threshold,
                layer2_enabled, cleared_addresses,
            )

        if net_error_seen:
            st.warning("Couldn't reach the pincode API for some rows (no internet or API down). "
                       "Those rows were checked with Layers 1 & 3 only.")

        flagged_df = result_df[result_df["Severity"].isin(["Critical", "Warning"])].reset_index(drop=True)
        clean_count = len(result_df) - len(flagged_df)

        st.subheader("3. Results")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total records", len(result_df))
        m2.metric("Critical", int((result_df["Severity"] == "Critical").sum()))
        m3.metric("Warning", int((result_df["Severity"] == "Warning").sum()))
        m4.metric("Clean", clean_count)

        if len(flagged_df) > 0:
            issue_counts = {}
            for issues in flagged_df["Issues"]:
                for i in issues.split("; "):
                    if not i:
                        continue
                    base = i.split("(")[0]
                    issue_counts[base] = issue_counts.get(base, 0) + 1
            st.bar_chart(pd.Series(issue_counts, name="Count"))

        # ---------------- Layer 4: AI review ----------------
        if use_ai and api_key and len(result_df) > 0:
            if ai_scope == "Flagged rows only":
                target_df = flagged_df
            elif ai_scope == "Random sample":
                target_df = result_df.sample(min(int(sample_size), len(result_df)))
            else:
                target_df = result_df
            ai_targets = target_df.index.tolist()

            if len(ai_targets) > 0:
                st.subheader("🤖 Layer 4: AI review")
                progress = st.progress(0.0, text="Calling Gemini...")
                ai_results = {}
                for n, idx in enumerate(ai_targets):
                    addr = target_df.loc[idx, "Address"]
                    if isinstance(addr, str) and addr.strip():
                        res = gemini_check(addr, api_key)
                    else:
                        res = {"verdict": "issue", "severity": "Critical", "category": "empty", "reason": "empty address"}
                    ai_results[idx] = res
                    progress.progress((n + 1) / len(ai_targets), text=f"Calling Gemini... {n+1}/{len(ai_targets)}")
                    time.sleep(1.1)  # pacing for free-tier rate limits
                progress.empty()

                ai_df = pd.DataFrame.from_dict(ai_results, orient="index")
                ai_df.columns = [f"AI_{c}" for c in ai_df.columns]
                result_df = result_df.join(ai_df)
                flagged_df = result_df[
                    result_df["Severity"].isin(["Critical", "Warning"]) |
                    (result_df.get("AI_verdict") == "issue")
                ].reset_index(drop=True)

        st.markdown("**Flagged addresses**")
        st.dataframe(flagged_df, use_container_width=True)

        st.markdown("**Full results**")
        st.dataframe(result_df, use_container_width=True)

        st.session_state["last_flagged_df"] = flagged_df
        st.session_state["last_result_df"] = result_df

        def to_excel_bytes(d):
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                d.to_excel(writer, index=False, sheet_name="Results")
            return buf.getvalue()

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button("⬇ Download flagged only (.xlsx)", data=to_excel_bytes(flagged_df),
                                file_name="flagged_addresses.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with dl2:
            st.download_button("⬇ Download full results (.xlsx)", data=to_excel_bytes(result_df),
                                file_name="all_address_results.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ---------------- Review queue ----------------
    if "last_flagged_df" in st.session_state and len(st.session_state["last_flagged_df"]) > 0:
        st.subheader("4. Review queue (human-in-the-loop)")
        st.caption("Mark false positives here. They're remembered on future runs so you "
                   "don't have to re-review the same address twice.")

        review_input = st.session_state["last_flagged_df"][["Agreement No", "Address", "Severity", "Issue Details"]].copy()
        review_input["Decision"] = "Pending"
        edited = st.data_editor(
            review_input,
            column_config={
                "Decision": st.column_config.SelectboxColumn(
                    options=["Pending", "Confirmed Issue", "False Positive"]
                )
            },
            use_container_width=True,
            key="review_editor",
        )

        if st.button("💾 Save reviewer decisions"):
            decided = edited[edited["Decision"] != "Pending"].copy()
            decided["Notes"] = ""
            new_feedback = decided[["Agreement No", "Address", "Decision", "Notes"]]
            combined = pd.concat([st.session_state.feedback_df, new_feedback], ignore_index=True)
            combined = combined.drop_duplicates(subset=["Address", "Decision"], keep="last")
            st.session_state.feedback_df = combined
            save_feedback(combined)
            st.success(f"Saved {len(new_feedback)} decisions. They'll auto-apply on future runs.")
else:
    st.info("Upload an Excel file above, or click 'Try demo data' to see it in action.")
