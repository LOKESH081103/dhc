"""
Cash / Cheque / DD — TAT Validation.

Standalone project, separate from the main DHC Working Automation app.
Same DCR source file, but a narrower job: flag CASH receipts not
deposited (seal challan uploaded) within 1 working day, and CHEQUE/DD
receipts not deposited within 3 working days.

PENDING DAYS = "Date" (column A, the report date) − "RECEIPT ENTER
DATE" (column E), both taken as calendar dates (time-of-day ignored).
This is a plain calendar-day count (no weekend/holiday skipping).

DATE PARSING NOTE: the file is read with engine="pyxlsb", and pyxlsb
cannot distinguish an Excel date cell from a plain number (pandas'
pyxlsb reader hands back the raw serial number either way). So "Date"
and "RECEIPT ENTER DATE" arrive here as numbers like 46213, not
datetimes. _parse_dates() below converts those serials back to real
dates (Excel's day-zero is 1899-12-30). It also transparently handles
the case where a column instead comes through as text like
"03-Jul-26" (dayfirst), so this keeps working if the reader ever
changes.
"""
import html
import re
from datetime import datetime

import pandas as pd

MODE_TAT = {"Cash": 1, "Cheque": 3, "DD": 3}
MODE_MATCH = {
    "Cash": ["CASH"],
    "Cheque": ["CHEQUE", "CHQ"],
    "DD": ["DD"],
}

# (display column name in the output/mail table, candidate raw-file column names to look for)
OUTPUT_COLUMNS = [
    ("RECEIPT ENTER DATE", ["RECEIPT ENTER DATE"]),
    ("Receipt No", ["Receipt No", "RECEIPT NO", "RECEIPTNO", "Receipt Number"]),
    ("AGREEMENTNO", ["AGREEMENTNO", "AGREEMENT NO", "Agreement No"]),
    ("PAYERNAME", ["PAYERNAME", "PAYER NAME"]),
    ("AMOUNTPAID", ["AMOUNTPAID", "AMOUNT PAID"]),
    ("COLLECTIONAGENTID", ["COLLECTIONAGENTID", "COLLECTION AGENT ID"]),
    ("COLLECTIONAGENTNAME", ["COLLECTIONAGENTNAME", "COLLECTION AGENT NAME"]),
    ("MODEOFPAYMENT", ["MODEOFPAYMENT", "MODE OF PAYMENT"]),
    ("RECEIPTTYPE", ["RECEIPTTYPE", "RECEIPT TYPE"]),
    ("BRANCH NAME", ["BRANCH NAME", "BRANCHNAME"]),
    ("NEW AREA", ["NEW AREA", "AREA"]),
    ("SUB REGION", ["SUB REGION", "Sub Region"]),
    ("MAIN REGION", ["MAIN REGION", "Region"]),
    ("Sub Zone", ["Sub Zone", "SUB ZONE"]),
    ("ZONE NEW", ["ZONE NEW", "ZONE", "Zone"]),
    ("SLAB", ["SLAB", "Slab"]),
]
DISPLAY_COLUMNS = ["PENDING DAYS"] + [c for c, _ in OUTPUT_COLUMNS]


def _normalize(name: str) -> str:
    return re.sub(r"\s+", " ", str(name).strip()).upper()


def _find_col(df: pd.DataFrame, *candidates: str) -> str | None:
    """Case/whitespace-insensitive column lookup, with a loose contains-match fallback."""
    norm_map = {_normalize(c): c for c in df.columns}
    for cand in candidates:
        key = _normalize(cand)
        if key in norm_map:
            return norm_map[key]
    for cand in candidates:
        key = _normalize(cand)
        for norm, orig in norm_map.items():
            if key in norm or norm in key:
                return orig
    return None


def _parse_dates(series: pd.Series) -> pd.Series:
    """
    Converts a column that may be Excel serial numbers (what pyxlsb
    gives us for date cells) and/or date text (e.g. "03-Jul-26") into
    real dates, normalized to midnight. Numeric values win when
    present; text is parsed with dayfirst=True as a fallback for any
    cell that isn't numeric.
    """
    numeric = pd.to_numeric(series, errors="coerce")
    from_serial = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    from_text = pd.to_datetime(series, errors="coerce", dayfirst=True)
    result = from_serial.where(numeric.notna(), from_text)
    return result.dt.normalize()


def load_dcr_raw(file) -> pd.DataFrame:
    """Reads Sheet1 of the uploaded DCR .xlsb, exactly as-is (no ETL transforms)."""
    df = pd.read_excel(file, sheet_name="Sheet1", engine="pyxlsb")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def filter_mode_exceeding(df: pd.DataFrame, mode: str) -> tuple[pd.DataFrame, dict]:
    """
    Filters df to the selected payment mode, computes PENDING DAYS, and
    returns only rows exceeding that mode's TAT — sorted worst (most
    overdue) first — plus a `meta` dict describing what was used/found
    so the app can show that transparently (present-date source, any
    columns it couldn't find, counts).
    """
    mop_col = _find_col(df, "MODEOFPAYMENT", "MODE OF PAYMENT")
    if mop_col is None:
        raise ValueError("Couldn't find a MODEOFPAYMENT column in this file.")
    receipt_enter_col = _find_col(df, "RECEIPT ENTER DATE")
    if receipt_enter_col is None:
        raise ValueError("Couldn't find a RECEIPT ENTER DATE column in this file.")

    # "Date" is always column A in this file — the report date, used as
    # the reference point for PENDING DAYS. Sanity-check its name (case/
    # whitespace-insensitive) rather than blindly trusting position.
    date_col = df.columns[0]
    if _normalize(date_col) != "DATE":
        found = _find_col(df, "Date")
        if found is not None:
            date_col = found

    match_values = [v.upper() for v in MODE_MATCH[mode]]
    mask = df[mop_col].astype(str).str.strip().str.upper().isin(match_values)
    sub = df.loc[mask].copy()
    total_matched = int(len(sub))

    receipt_norm = _parse_dates(sub[receipt_enter_col])
    present_norm = _parse_dates(sub[date_col])
    present_source = f'the "{date_col}" column (report date)'

    # Drop rows where either date failed to parse — can't compute TAT for them.
    valid = receipt_norm.notna() & present_norm.notna()
    unparseable = int((~valid).sum())
    sub = sub.loc[valid].copy()
    receipt_norm = receipt_norm.loc[valid]
    present_norm = present_norm.loc[valid]

    sub["_PENDING_DAYS"] = (present_norm - receipt_norm).dt.days
    sub[receipt_enter_col] = receipt_norm  # normalized, for clean display below

    threshold = MODE_TAT[mode]
    exceeding = sub[sub["_PENDING_DAYS"] > threshold].copy()
    exceeding = exceeding.sort_values(["_PENDING_DAYS", receipt_enter_col], ascending=[False, True])

    missing: list[str] = []
    out = pd.DataFrame(index=exceeding.index)
    out["PENDING DAYS"] = exceeding["_PENDING_DAYS"].astype(int)
    for display_name, candidates in OUTPUT_COLUMNS:
        col = _find_col(df, *candidates)
        if col is None:
            missing.append(display_name)
            out[display_name] = "N/A"
        else:
            out[display_name] = exceeding[col]

    out["RECEIPT ENTER DATE"] = pd.to_datetime(out["RECEIPT ENTER DATE"], errors="coerce").dt.strftime("%d-%b-%y")
    out = out.reset_index(drop=True)

    meta = {
        "mode": mode,
        "threshold": threshold,
        "present_source": present_source,
        "missing_columns": missing,
        "total_matched": total_matched,
        "total_exceeding": int(len(out)),
        "unparseable_dates": unparseable,
    }
    return out, meta


def _mode_wording(mode: str) -> dict:
    """
    Cash wording below is your exact provided text, verbatim. Cheque/DD
    wording is my best-effort parallel — you gave the TAT (3 days) and
    the rule ("exceeding it, it should be displayed") but not an exact
    template for those two, so please check this paragraph reads right
    before relying on it; happy to adjust the exact phrasing.
    """
    if mode == "Cash":
        return {
            "item": "Cash Payments",
            "intro": (
                "Please find the below mentioned Cash Payments that have been delayed in deposit "
                "/uploading the seal challan in the system for more than one working day."
            ),
            "note1": (
                "As per Audit Concerns, it is mandatory that any cash payment is deposited and the "
                "corresponding seal challan is uploaded into the system within one working day, and "
                "these cases will be sent to FCU by Operations for review, so please act on top priority"
            ),
        }
    item = f"{mode} Payments"
    return {
        "item": item,
        "intro": (
            f"Please find the below mentioned {item} that have been delayed in deposit / updation "
            "in the system for more than three working days."
        ),
        "note1": (
            f"As per Audit Concerns, it is mandatory that any {mode.lower()} payment is deposited and "
            "updated into the system within three working days, and these cases will be sent to FCU by "
            "Operations for review, so please act on top priority"
        ),
    }


NOTE2 = "Kindly ensure that we comply with this guideline moving forward to avoid any audit discrepancies."
SIGNOFF = "Regards,\nShalini R\nLAP Collection Support \u2013 Chennai HO\n9600145482"


def build_tat_mail_html(mode: str, exceeding_df: pd.DataFrame) -> tuple[str, str]:
    """(subject, HTML body) for the Outlook "Compose in Outlook" button — full table, every row."""
    e = html.escape
    w = _mode_wording(mode)
    subject = f"{w['item']} Delayed Beyond TAT \u2014 {datetime.now().strftime('%d-%b-%Y')}"

    header_html = "".join(
        f'<th style="padding:6px 10px;border:1px solid #d1d5db;background:#1e293b;'
        f'color:#ffffff;font-size:12px;text-align:left;white-space:nowrap;">{e(c)}</th>'
        for c in DISPLAY_COLUMNS
    )
    rows_html = ""
    for i, (_, row) in enumerate(exceeding_df.iterrows()):
        bg = "#ffffff" if i % 2 == 0 else "#f3f4f6"
        cells = "".join(
            f'<td style="padding:5px 10px;border:1px solid #e5e7eb;font-size:12px;'
            f'color:#111827;white-space:nowrap;">{e(str(row[c]))}</td>'
            for c in DISPLAY_COLUMNS
        )
        rows_html += f'<tr style="background:{bg};">{cells}</tr>'

    table_html = (
        '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;'
        f'font-family:Segoe UI,Arial,sans-serif;"><tr>{header_html}</tr>{rows_html}</table>'
    )

    body = f"""
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#111827;line-height:1.6;">
      <p>Dear Team,</p>
      <p>{e(w['intro'])}</p>
      <p>Note: {e(w['note1'])}<br/>{e(NOTE2)}</p>
      {table_html}
      <p style="margin-top:20px;">{e(SIGNOFF).replace(chr(10), '<br/>')}</p>
    </div>
    """
    return subject, body


def build_tat_mail_text(mode: str, exceeding_df: pd.DataFrame, max_rows: int = 20) -> tuple[str, str]:
    """(subject, plain text) for the mailto fallback button. Capped at max_rows — mailto: has a hard length limit most mail clients enforce, unlike the Outlook COM path above."""
    w = _mode_wording(mode)
    subject = f"{w['item']} Delayed Beyond TAT \u2014 {datetime.now().strftime('%d-%b-%Y')}"

    lines = ["Dear Team,", "", w["intro"], "", f"Note: {w['note1']}", NOTE2, ""]
    shown = exceeding_df.head(max_rows)
    for _, row in shown.iterrows():
        lines.append(
            f"- {row['Receipt No']} | {row['RECEIPT ENTER DATE']} | Pending {row['PENDING DAYS']}d | "
            f"{row['PAYERNAME']} | Rs {row['AMOUNTPAID']}"
        )
    if len(exceeding_df) > max_rows:
        lines.append(f"... and {len(exceeding_df) - max_rows} more (use 'Compose in Outlook' for the full list).")
    lines += ["", SIGNOFF]
    return subject, "\n".join(lines)


def compose_outlook_mail_html(subject: str, html_body: str) -> None:
    """
    Opens (does NOT send) a new mail draft directly in classic Outlook
    via COM automation. Requires pywin32 (`pip install pywin32`) and a
    Windows machine with classic Outlook installed.

    Calls pythoncom.CoInitialize() first: Streamlit runs each button's
    callback on a worker thread (not the main thread), and COM requires
    the calling thread to be initialized before any COM object is
    created on it.
    """
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Display()
    finally:
        pythoncom.CoUninitialize()
