"""
Cash / Cheque / DD — TAT Validation & Collection Dashboards.

Standalone project, separate from the main DHC Working Automation app.
Same DCR source file, but a narrower job: flag CASH receipts not
deposited (seal challan uploaded) within 1 working day, and CHEQUE/DD
receipts not deposited within 3 working days.
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
    Filters df to the selected payment mode and "Updation Pending" status, 
    computes PENDING DAYS, and returns only rows exceeding that mode's TAT 
    — sorted worst (most overdue) first.
    """
    mop_col = _find_col(df, "MODEOFPAYMENT", "MODE OF PAYMENT")
    if mop_col is None:
        raise ValueError("Couldn't find a MODEOFPAYMENT column in this file.")
        
    receipt_enter_col = _find_col(df, "RECEIPT ENTER DATE")
    if receipt_enter_col is None:
        raise ValueError("Couldn't find a RECEIPT ENTER DATE column in this file.")

    status_col = _find_col(df, "RECEIPT STATUS", "RECEIPTSTATUS", "STATUS")
    if status_col is None:
        raise ValueError("Couldn't find a RECEIPT STATUS column in this file.")

    date_col = df.columns[0]
    if _normalize(date_col) != "DATE":
        found = _find_col(df, "Date")
        if found is not None:
            date_col = found

    match_values = [v.upper() for v in MODE_MATCH[mode]]
    mode_mask = df[mop_col].astype(str).str.strip().str.upper().isin(match_values)
    
    status_mask = df[status_col].astype(str).str.strip().str.upper() == "UPDATION PENDING"
    
    sub = df.loc[mode_mask & status_mask].copy()
    total_matched = int(len(sub))

    receipt_norm = _parse_dates(sub[receipt_enter_col])
    present_norm = _parse_dates(sub[date_col])
    present_source = f'the "{date_col}" column (report date)'

    valid = receipt_norm.notna() & present_norm.notna()
    unparseable = int((~valid).sum())
    sub = sub.loc[valid].copy()
    receipt_norm = receipt_norm.loc[valid]
    present_norm = present_norm.loc[valid]

    sub["_PENDING_DAYS"] = (present_norm - receipt_norm).dt.days
    sub[receipt_enter_col] = receipt_norm

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
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Display()
    finally:
        pythoncom.CoUninitialize()

# ==========================================
# NEW FEATURE: AIRTEL GATEWAY DASHBOARD
# ==========================================
# ==========================================
# NEW FEATURE: AIRTEL GATEWAY DASHBOARD
# ==========================================
# ==========================================
# NEW FEATURE: AIRTEL GATEWAY DASHBOARD
# ==========================================
# ==========================================
# NEW FEATURE: AIRTEL GATEWAY DASHBOARD
# ==========================================
def _build_pivot(df: pd.DataFrame, index_cols: list, index_names: list) -> pd.DataFrame:
    """Helper function to build a MultiIndex pivot table with Counts, Values, and Percentages."""
    pivot = pd.pivot_table(
        df,
        index=index_cols,
        columns='Dash_MOP',
        values='Amount_Num',
        aggfunc=['count', 'sum'],
        fill_value=0
    )
    
    data = {}
    for mode in ['Deposited - Airtel', 'Deposited - Bank']:
        if 'count' in pivot.columns and mode in pivot['count'].columns:
            data[(mode, 'Count')] = pivot['count'][mode]
            data[(mode, 'Value')] = pivot['sum'][mode] / 100000.0
        else:
            data[(mode, 'Count')] = pd.Series(0, index=pivot.index)
            data[(mode, 'Value')] = pd.Series(0.0, index=pivot.index)
            
    new_df = pd.DataFrame(data, index=pivot.index)
    
    new_df[('Overall', 'Count')] = new_df[('Deposited - Airtel', 'Count')] + new_df[('Deposited - Bank', 'Count')]
    new_df[('Overall', 'Value')] = new_df[('Deposited - Airtel', 'Value')] + new_df[('Deposited - Bank', 'Value')]
    
    for mode in ['Deposited - Airtel', 'Deposited - Bank']:
        new_df[(mode, 'Count %')] = (new_df[(mode, 'Count')] / new_df[('Overall', 'Count')].replace(0, 1)) * 100
        new_df[(mode, 'Value %')] = (new_df[(mode, 'Value')] / new_df[('Overall', 'Value')].replace(0, 1)) * 100
        
    ordered_cols = [
        ('Deposited - Airtel', 'Count'), ('Deposited - Airtel', 'Value'), 
        ('Deposited - Airtel', 'Count %'), ('Deposited - Airtel', 'Value %'),
        ('Deposited - Bank', 'Count'), ('Deposited - Bank', 'Value'), 
        ('Deposited - Bank', 'Count %'), ('Deposited - Bank', 'Value %'),
        ('Overall', 'Count'), ('Overall', 'Value')
    ]
    new_df = new_df[ordered_cols]
    
    # Sort alphabetically (E -> N -> S -> W) so it naturally orders East, North, South, West
    new_df = new_df.sort_index()
    
    # Handle Grand Total for Multi-level index (Zone -> Region) vs Single-level (Zone)
    if len(index_cols) > 1:
        gt_idx = tuple(['Grand Total'] + [''] * (len(index_cols) - 1))
    else:
        gt_idx = 'Grand Total'
        
    new_df.loc[gt_idx, :] = new_df.sum()
    
    # Recalculate Grand Total percentages
    gt = gt_idx
    for mode in ['Deposited - Airtel', 'Deposited - Bank']:
        gt_count_pct = (new_df.loc[gt, (mode, 'Count')] / new_df.loc[gt, ('Overall', 'Count')] * 100) if new_df.loc[gt, ('Overall', 'Count')] > 0 else 0
        gt_value_pct = (new_df.loc[gt, (mode, 'Value')] / new_df.loc[gt, ('Overall', 'Value')] * 100) if new_df.loc[gt, ('Overall', 'Value')] > 0 else 0
        new_df.loc[gt, (mode, 'Count %')] = gt_count_pct
        new_df.loc[gt, (mode, 'Value %')] = gt_value_pct
    
    new_df.index.names = index_names
    return new_df.reset_index()


def generate_airtel_dashboard(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    zone_col = _find_col(df, "ZONE NEW", "ZONE", "Zone")
    region_col = _find_col(df, "MAIN REGION", "Region", "REGION")
    mop_col = _find_col(df, "MODEOFPAYMENT", "MODE OF PAYMENT")
    amount_col = _find_col(df, "AMOUNTPAID", "AMOUNT PAID")

    if not all([zone_col, region_col, mop_col, amount_col]):
        raise ValueError("Missing required columns for Dashboard (Zone, Region, MOP, or Amount).")

    df_dash = df.copy()
    df_dash['Amount_Num'] = pd.to_numeric(df_dash[amount_col], errors='coerce').fillna(0)

    def categorize_mop(m):
        m_upper = str(m).strip().upper()
        if 'AIRTEL' in m_upper:
            return 'Deposited - Airtel'
        elif 'CASH' in m_upper:
            return 'Deposited - Bank'
        return 'Other'

    df_dash['Dash_MOP'] = df_dash[mop_col].apply(categorize_mop)
    df_filtered = df_dash[df_dash['Dash_MOP'].isin(['Deposited - Airtel', 'Deposited - Bank'])]

    # Group by Zone ONLY
    zone_summary = _build_pivot(df_filtered, [zone_col], ["Zone"])
    # Group by Zone THEN Region (creates the hierarchical layout)
    region_summary = _build_pivot(df_filtered, [zone_col, region_col], ["Zone", "Region"])

    return zone_summary, region_summary


def build_airtel_mail_html(zone_df: pd.DataFrame, region_df: pd.DataFrame, bank_pct: float) -> tuple[str, str]:
    subject = f"Airtel Money Gateway Collection \u2014 {datetime.now().strftime('%d-%b-%Y')}"
    
    def _df_to_styled_html(df):
        formatters = {}
        for col in df.columns:
            if isinstance(col, tuple):
                if col[1] in ['Count %', 'Value %']:
                    formatters[col] = lambda x: f"{x:,.2f}%"
                elif col[1] == 'Count':
                    formatters[col] = lambda x: f"{x:,.0f}"
                elif col[1] == 'Value':
                    formatters[col] = lambda x: f"{x:,.2f}"
                
        table_html = df.to_html(index=False, formatters=formatters, na_rep="0", justify='center')
        table_html = table_html.replace('<table border="1" class="dataframe">', '<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:12px;text-align:center;width:90%;margin-bottom:20px;" border="1">')
        table_html = table_html.replace('<th>', '<th style="background-color:#1e293b;color:#ffffff;padding:6px;border:1px solid #d1d5db;">')
        table_html = table_html.replace('<td>', '<td style="padding:5px;border:1px solid #d1d5db;color:#111827;">')
        return table_html

    body = f"""
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#111827;line-height:1.6;">
      <p>Dear Team,</p>
      <p>Please find the below Zone / region wise Cash Collection done by CFEs through CCP Mobile Application.<br>
      Deposition is Airtel Outlet Vs Bank details is shown below.<br>
      Ensure that 100% cash is being deposited only in Airtel payment bank.</p>
      
      <p><b>{bank_pct:.2f}%</b> of the Cash got Deposited in Bank, Please reduce the same.</p>
      <br>
      <b>Zone-wise Dashboard</b><br>
      {_df_to_styled_html(zone_df)}
      <br>
      <b>Region-wise Dashboard</b><br>
      {_df_to_styled_html(region_df)}
      <p style="margin-top:20px;">{html.escape(SIGNOFF).replace(chr(10), '<br/>')}</p>
    </div>
    """
    return subject, body