import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import io
import re
import collections

# Page Configuration
st.set_page_config(
    page_title="General Ledger Auditor & Auto-Fixer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1F4E79;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #555555;
        margin-bottom: 1.5rem;
    }
    .upload-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 1.5rem;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        margin-bottom: 1rem;
    }
    .step-badge {
        background-color: #1F4E79;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-weight: 700;
        font-size: 0.85rem;
        margin-right: 6px;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# CATEGORY COMPATIBILITY & EQUIVALENCE GROUPS
# -----------------------------------------------------------------------------

EQUIVALENCE_GROUPS = [
    # 1. Client Revenue Cycle (A/R, Services, Invoicing, Collections)
    {"accounts receivable (a/r)", "services", "accounts receivable", "services income", "income", "sales", "revenue"},
    
    # 2. Subcontractor / Direct Vendor Cycle (A/P, COGS - Vendor, Consulting Fees)
    {"accounts payable (a/p)", "cogs - vendor", "consulting fees", "accounts payable", "cost of goods sold", "vendor cogs", "cogs"},
    
    # 3. Owner Distributions / Draws
    {"owners distribution", "shareholders' equity:distributions", "distributions", "shareholders' equity", "owner's distribution"},
    
    # 4. Payroll Processing
    {"payroll expenses:payroll processing fee", "payroll processing fee", "payroll processing fees"},
    
    # 5. Payroll Taxes
    {"payroll wages and tax to pay:payroll tax to pay", "payroll tax to pay", "payroll expenses:payroll taxes", "payroll taxes"},
    
    # 6. Payroll Wages & Salaries
    {"payroll wages payable", "salaries & wages", "cogs - vendor:salaries & wages", "payroll wages"},
    
    # 7. Rent & Facilities
    {"rent:building & land rent", "building & land rent", "rent"},
]

def normalize_cat(cat):
    return str(cat).strip().lower()

def are_categories_compatible(cat1, cat2):
    """Check if two category strings are legitimate accounting counterparts or sub-accounts."""
    c1 = normalize_cat(cat1)
    c2 = normalize_cat(cat2)
    
    if c1 == c2:
        return True
        
    # Sub-account matching (e.g. "phone service" == "utilities:phone service")
    if c1.endswith(":" + c2) or c2.endswith(":" + c1):
        return True
    if ":" in c1 and ":" not in c2 and c1.split(":")[-1] == c2:
        return True
    if ":" in c2 and ":" not in c1 and c2.split(":")[-1] == c1:
        return True
        
    # Equivalence groups (e.g. Services vs Accounts Receivable (A/R))
    for grp in EQUIVALENCE_GROUPS:
        c1_in = any(c1 == m or c1.endswith(":" + m) or m.endswith(":" + c1) for m in grp)
        c2_in = any(c2 == m or c2.endswith(":" + m) or m.endswith(":" + c2) for m in grp)
        if c1_in and c2_in:
            return True
            
    return False

def clean_vendor(name, desc):
    """Normalize and extract merchant/vendor identity from Name or Description."""
    text = str(name).strip() if name and str(name).strip() and str(name).strip() != "None" else (str(desc).strip() if desc else "")
    if not text:
        return "Unknown"
    
    t_upper = text.upper()
    
    # Credit Cards
    if "AMERICAN EXPRESS" in t_upper or "AMEX" in t_upper:
        return "Amex CC"
    if "BANK OF AMERICA - CREDIT CARD" in t_upper or "BOA CC" in t_upper or "B OF A CREDIT CARD" in t_upper:
        return "BOA CC"
    # ADP Payroll
    if "ADP PAYROLL FEES" in t_upper or "ADP FEES" in t_upper:
        return "ADP Payroll Processing"
    if "ADP TAX" in t_upper:
        return "ADP Payroll Tax"
    if "ADP 401K" in t_upper:
        return "ADP 401k"
    if "ADP WAGE" in t_upper:
        return "ADP Wage Pay"
    # Food & Delivery
    if "DOORDASH" in t_upper or "DD *" in t_upper:
        if "TARGET" in t_upper or "ACE HARDWARE" in t_upper:
            return "DoorDash - Retail Supplies"
        return "DoorDash"
    if "UBER EATS" in t_upper:
        return "Uber Eats"
    if "UBER" in t_upper:
        return "Uber"
    # Tech & Software
    if "OPENAI" in t_upper or "CHATGPT" in t_upper:
        return "OpenAI"
    if "ANTHROPIC" in t_upper or "CLAUDE" in t_upper:
        return "Anthropic"
    if "LINKEDIN" in t_upper:
        return "LinkedIn"
    if "GOOGLE WORKSPACE" in t_upper or "GOOGLE" in t_upper:
        return "Google Workspace"
    if "MICROSOFT" in t_upper:
        return "Microsoft"
    if "AMAZON" in t_upper:
        return "Amazon"
    if "INTUIT" in t_upper or "QUICKBOOKS" in t_upper:
        return "Intuit QuickBooks"
    if "NETFLIX" in t_upper:
        return "Netflix"
    if "PRIME VIDEO" in t_upper:
        return "Prime Video"
    if "HP INSTANT INK" in t_upper:
        return "HP Instant Ink"
    # Facilities, Travel & Fuel
    if "BANNER PEST" in t_upper:
        return "Banner Pest Services"
    if "QUIK STOP" in t_upper:
        return "Quik Stop"
    if "UNION 76" in t_upper:
        return "Union 76"
    if "CHEVRON" in t_upper:
        return "Chevron"
    if "WAWA" in t_upper:
        return "Wawa"
    if "FASTRAK" in t_upper:
        return "Fastrak"
    if "SPOTHERO" in t_upper:
        return "SpotHero"
    # Telecom & Utilities
    if "DELMARVA" in t_upper:
        return "Delmarva Power"
    if "COMCAST" in t_upper or "XFINITY" in t_upper:
        return "Comcast / Xfinity"
    if "AT&T" in t_upper:
        return "AT&T"
    if "VERIZON" in t_upper:
        return "Verizon"
    if "RINGCENTRAL" in t_upper:
        return "RingCentral"
    if "VONAGE" in t_upper:
        return "Vonage"
    # Legal, Insurance & Subcontractors
    if "CSAA" in t_upper:
        return "CSAA Insurance"
    if "GLOBAL IMMIGRATION" in t_upper or "GLOBALIMMIGRATIO" in t_upper:
        return "Global Immigration Partners"
    if "SYNERGY BADMINTON" in t_upper:
        return "Synergy Badminton"
    if "FREEWHEEL BREWING" in t_upper:
        return "Freewheel Brewing"
    if "ACTIVATE PLANO" in t_upper:
        return "Activate Plano"
    if "CLUB 28 BADMINTON" in t_upper:
        return "Club 28 Badminton"
    if "ASTRAMIND" in t_upper:
        return "Astramind Solutions"
    if "ANERU" in t_upper:
        return "Aneru LLC"
    if "PTL CONSULTANCY" in t_upper:
        return "PTL Consultancy"
    if "SYSCONSULT" in t_upper:
        return "Sysconsult LLC"
    if "SAGE SOLUTIONS" in t_upper:
        return "Sage Solutions"
    if "AMERICAN TECHNOLOGY SERVICES" in t_upper:
        return "American Technology Services LLC"
    if "AGILISIUM" in t_upper:
        return "Agilisium Consulting LLC"
    if "MEDIX" in t_upper:
        return "Medix Staffing Solutions"
    if "NITYO" in t_upper:
        return "Nityo Infotech Corporation"
    if "INFINITE COMPUTER" in t_upper:
        return "Infinite Computer Solutions"
    if "PHOENIX STAFF" in t_upper:
        return "Phoenix Staff Inc"
    if "BIGBYTE" in t_upper:
        return "Bigbyte"
    if "SANA SOFTWARE" in t_upper:
        return "SANA Software Solutions"
    if "VIVID TECHNOLOGIES" in t_upper:
        return "Vivid Technologies Inc"
    if "VENDORPASS" in t_upper:
        return "Vendorpass Inc"
    if "RANDSTAND" in t_upper or "RANDSTAD" in t_upper:
        return "Randstad Digital"
    if "SANVI" in t_upper:
        return "Sanvi Consulting Services LLC"
    
    clean = re.sub(r"^(AplPay|TST\*\s*|DD\s*\*\s*|PY\s*\*\s*|PAYPAL\s*\*|SQ\s*\*)", "", text, flags=re.IGNORECASE).strip()
    return clean[:35]

def extract_historical_rules(hist_wb):
    """Dynamically learn rules from historical workbook."""
    sheet = hist_wb.active
    rows = list(sheet.iter_rows(values_only=True))
    learned = collections.defaultdict(collections.Counter)
    
    for r in rows:
        c_dist = str(r[1]).strip() if r[1] is not None else ""
        c_name = str(r[5]).strip() if r[5] is not None else ""
        c_desc = str(r[6]).strip() if r[6] is not None else ""
        c_split = str(r[7]).strip() if r[7] is not None else ""
        
        if not c_dist or c_dist in ["Beginning Balance", "Distribution account"]:
            continue
            
        category = c_split if c_dist in ["Bank of America", "Amex CC", "BOA CC"] else c_dist
        vendor = clean_vendor(c_name, c_desc)
        
        if vendor and vendor != "Unknown" and category and "Uncategorized" not in category:
            learned[vendor][category] += 1
            
    rules = {}
    for v, cat_counts in learned.items():
        top_cat = cat_counts.most_common(1)[0][0]
        rules[v] = {
            "primary_category": top_cat,
            "occurrences": sum(cat_counts.values()),
            "all_categories": dict(cat_counts)
        }
    return rules

def audit_and_fix_current_ledger(curr_wb, historical_rules):
    """Audit current ledger, auto-fix credit card vendor names, match historical categories, and style all rows."""
    sheet = curr_wb.active
    
    # Styles
    header_bg = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_cc_updated = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")      # Blue
    fill_mismatch = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")        # Yellow
    fill_uncategorized = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")   # Red
    fill_matched = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid")         # Green
    
    font_bold = Font(name="Arial", size=9, bold=True)
    font_regular = Font(name="Arial", size=9)
    font_danger = Font(name="Arial", size=9, color="721C24", bold=True)
    font_warning = Font(name="Arial", size=9, color="856404", bold=True)
    font_info = Font(name="Arial", size=9, color="0C5460", bold=True)
    font_success = Font(name="Arial", size=9, color="155724")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    cell_k5 = sheet.cell(row=5, column=11, value="Audit Status")
    cell_l5 = sheet.cell(row=5, column=12, value="Audit Details & Suggested Category")
    for cell in [cell_k5, cell_l5]:
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = header_bg
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    audit_results = []
    cc_fixed = 0
    mismatch_count = 0
    uncat_count = 0
    matched_count = 0
    
    for r in range(6, sheet.max_row + 1):
        c_dist = sheet.cell(row=r, column=2).value
        c_date = sheet.cell(row=r, column=3).value
        c_type = sheet.cell(row=r, column=4).value
        c_num = sheet.cell(row=r, column=5).value
        cell_name = sheet.cell(row=r, column=6)
        c_name = cell_name.value
        c_desc = sheet.cell(row=r, column=7).value
        c_split = sheet.cell(row=r, column=8).value
        c_amt = sheet.cell(row=r, column=9).value
        
        if not c_dist or c_dist in ["Beginning Balance", "Distribution account"]:
            continue
            
        str_dist = str(c_dist).strip()
        str_type = str(c_type).strip() if c_type else ""
        str_name = str(c_name).strip() if c_name else ""
        str_desc = str(c_desc).strip() if c_desc else ""
        str_split = str(c_split).strip() if c_split else ""
        amt_val = float(c_amt) if c_amt is not None else 0.0
        
        # 1. Credit card payment vendor fix
        is_cc_payment = False
        target_cc_name = ""
        
        if str_type == "Credit Card Payment" or "CREDIT CARD Bill Payment" in str_desc or "AMERICAN EXPRESS DES:ACH PMT" in str_desc:
            is_cc_payment = True
            if "Amex CC" in str_dist or "Amex CC" in str_split or "AMERICAN EXPRESS" in str_desc:
                target_cc_name = "Amex CC"
            elif "BOA CC" in str_dist or "BOA CC" in str_split or "BANK OF AMERICA - CREDIT CARD" in str_desc:
                target_cc_name = "BOA CC"
            else:
                target_cc_name = "Credit Card"
        elif str_dist == "Bank of America" and str_split in ["Amex CC", "BOA CC"]:
            is_cc_payment = True
            target_cc_name = str_split
            
        if is_cc_payment:
            cell_name.value = target_cc_name
            status = "CC VENDOR UPDATED"
            note = f"Vendor Name auto-set to '{target_cc_name}'"
            row_fill = fill_cc_updated
            row_font = font_info
            cc_fixed += 1
            audit_results.append({
                "Row": r, "Date": c_date, "Account": str_dist, "Type": str_type,
                "Vendor": target_cc_name, "Description": str_desc, "Split": str_split,
                "Amount": amt_val, "Status": status, "Issue": "Missing CC Vendor Name",
                "Action": f"Set Vendor Name to '{target_cc_name}'"
            })
        else:
            # 2. Historical Category Matching
            vendor_key = clean_vendor(str_name, str_desc)
            
            if "Uncategorized" in str_dist or "Uncategorized" in str_split:
                status = "UNCATEGORIZED"
                note = "Uncategorized transaction requires reconciliation / invoice mapping."
                row_fill = fill_uncategorized
                row_font = font_danger
                uncat_count += 1
                audit_results.append({
                    "Row": r, "Date": c_date, "Account": str_dist, "Type": str_type,
                    "Vendor": str_name or vendor_key, "Description": str_desc, "Split": str_split,
                    "Amount": amt_val, "Status": status, "Issue": "Uncategorized Transaction",
                    "Action": "Reconcile with specific customer invoice or expense account"
                })
            elif vendor_key in historical_rules:
                rule_info = historical_rules[vendor_key]
                expected_category = rule_info["primary_category"]
                curr_category = str_split if str_dist in ["Bank of America", "Amex CC", "BOA CC"] else str_dist
                
                # Check compatibility with parent accounts and equivalence groups
                if not are_categories_compatible(curr_category, expected_category):
                    status = "CATEGORY MISMATCH"
                    note = f"Expected '{expected_category}' based on historical data, found '{curr_category}'."
                    row_fill = fill_mismatch
                    row_font = font_warning
                    mismatch_count += 1
                    audit_results.append({
                        "Row": r, "Date": c_date, "Account": str_dist, "Type": str_type,
                        "Vendor": str_name or vendor_key, "Description": str_desc, "Split": str_split,
                        "Amount": amt_val, "Status": status, "Issue": f"Mismatched Category (Found: {curr_category})",
                        "Action": f"Reclassify to historical category: '{expected_category}'"
                    })
                else:
                    status = "MATCHED"
                    note = f"Verified: {expected_category}"
                    row_fill = fill_matched
                    row_font = font_success
                    matched_count += 1
            else:
                status = "MATCHED"
                note = "Standard ledger entry"
                row_fill = fill_matched
                row_font = font_success
                matched_count += 1
                
        cell_k = sheet.cell(row=r, column=11, value=status)
        cell_l = sheet.cell(row=r, column=12, value=note)
        cell_k.font = row_font
        cell_l.font = font_regular
        
        for col_idx in range(2, 13):
            sheet.cell(row=r, column=col_idx).fill = row_fill
            
    if "Audit & Discrepancies" in curr_wb.sheetnames:
        del curr_wb["Audit & Discrepancies"]
    audit_ws = curr_wb.create_sheet(title="Audit & Discrepancies", index=0)
    audit_ws.views.sheetView[0].showGridLines = True
    
    audit_ws["A1"] = "General Ledger Audit & Historical Compliance Summary"
    audit_ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F4E79")
    audit_ws["A2"] = f"Total Transactions Audited: {len(audit_results) + matched_count} | Issues / Updates: {len(audit_results)}"
    audit_ws["A2"].font = Font(name="Arial", size=10, italic=True, color="595959")
    
    kpis = [
        ("Total Discrepancies & Flags", len(audit_results), "FFF3CD", "856404"),
        ("CC Vendor Payments Fixed", cc_fixed, "D1ECF1", "0C5460"),
        ("Category Inconsistencies", mismatch_count, "FFF3CD", "856404"),
        ("Uncategorized Items", uncat_count, "F8D7DA", "721C24"),
    ]
    col_start = 1
    for title, val, bg_hex, text_hex in kpis:
        c1 = audit_ws.cell(row=4, column=col_start, value=title)
        c2 = audit_ws.cell(row=5, column=col_start, value=val)
        c1.font = Font(name="Arial", size=9, bold=True, color=text_hex)
        c2.font = Font(name="Arial", size=16, bold=True, color=text_hex)
        c1.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
        c2.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
        c1.alignment = Alignment(horizontal="center")
        c2.alignment = Alignment(horizontal="center")
        audit_ws.merge_cells(start_row=4, start_column=col_start, end_row=4, end_column=col_start+1)
        audit_ws.merge_cells(start_row=5, start_column=col_start, end_row=5, end_column=col_start+1)
        col_start += 2
        
    headers = ["Row #", "Date", "Account", "Type", "Vendor / Name", "Description", "Amount", "Status", "Identified Issue", "Recommended Action"]
    for col_idx, h_text in enumerate(headers, 1):
        cell = audit_ws.cell(row=7, column=col_idx, value=h_text)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = header_bg
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for idx, item in enumerate(audit_results, 8):
        audit_ws.cell(row=idx, column=1, value=item["Row"]).alignment = Alignment(horizontal="center")
        audit_ws.cell(row=idx, column=2, value=item["Date"]).alignment = Alignment(horizontal="center")
        audit_ws.cell(row=idx, column=3, value=item["Account"])
        audit_ws.cell(row=idx, column=4, value=item["Type"])
        audit_ws.cell(row=idx, column=5, value=item["Vendor"]).font = font_bold
        audit_ws.cell(row=idx, column=6, value=item["Description"])
        c_amt = audit_ws.cell(row=idx, column=7, value=item["Amount"])
        c_amt.number_format = "$#,##0.00;($#,##0.00);\"-\""
        
        c_stat = audit_ws.cell(row=idx, column=8, value=item["Status"])
        if item["Status"] == "UNCATEGORIZED":
            c_stat.fill = fill_uncategorized
            c_stat.font = font_danger
        elif item["Status"] == "CATEGORY MISMATCH":
            c_stat.fill = fill_mismatch
            c_stat.font = font_warning
        else:
            c_stat.fill = fill_cc_updated
            c_stat.font = font_info
        c_stat.alignment = Alignment(horizontal="center")
        
        audit_ws.cell(row=idx, column=9, value=item["Issue"])
        audit_ws.cell(row=idx, column=10, value=item["Action"]).font = font_bold
        
        for c in range(1, 11):
            audit_ws.cell(row=idx, column=c).border = thin_border
            
    for ws in [audit_ws, sheet]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.row < 4 and ws == audit_ws:
                    continue
                if '\n' in val_str:
                    val_str = max(val_str.split('\n'), key=len)
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 42)
            
    summary_data = {
        "total_rows": len(audit_results) + matched_count,
        "cc_fixed": cc_fixed,
        "mismatch_count": mismatch_count,
        "uncat_count": uncat_count,
        "matched_count": matched_count,
        "discrepancies": audit_results
    }
    return curr_wb, summary_data


# -----------------------------------------------------------------------------
# STREAMLIT UI (DUAL UPLOAD)
# -----------------------------------------------------------------------------
st.markdown('<div class="main-header">📊 General Ledger Auditor & Auto-Fixer</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload your <b>Historical Ledger</b> (ground truth baseline) and your <b>Current Ledger</b> (month to audit & fix). The engine dynamically learns historical vendor patterns, fixes credit card vendor names, and color-codes all discrepancies.</div>', unsafe_allow_html=True)

# Sidebar
st.sidebar.header("🎨 Audit Color Legend")
st.sidebar.markdown("""
- 🟦 **Blue**: Credit Card Payment (Vendor Name Populated)
- 🟨 **Yellow**: Category Mismatch vs Past History
- 🟥 **Red**: Uncategorized Income / Expense
- 🟩 **Green**: Verified & Matched Entry
""")
st.sidebar.markdown("---")
st.sidebar.caption("All transactions are evaluated against the dynamic rules extracted from your historical ledger.")

# Dual Upload Section
col_hist, col_curr = st.columns(2)

with col_hist:
    st.markdown('<div class="upload-card"><span class="step-badge">STEP 1</span><b>Upload Historical Ledger (Ground Truth)</b>', unsafe_allow_html=True)
    hist_file = st.file_uploader(
        "Upload Past Correctly Classified Ledger (.xlsx or .xls)",
        type=["xlsx", "xls"],
        key="hist_uploader",
        help="Upload past general ledger data containing correct classifications (e.g. Jan-Jun)."
    )
    st.markdown('</div>', unsafe_allow_html=True)

with col_curr:
    st.markdown('<div class="upload-card"><span class="step-badge">STEP 2</span><b>Upload Current Ledger (To Audit & Fix)</b>', unsafe_allow_html=True)
    curr_file = st.file_uploader(
        "Upload Current Month Ledger (.xlsx or .xls)",
        type=["xlsx", "xls"],
        key="curr_uploader",
        help="Upload the new general ledger file to be audited, auto-fixed, and highlighted."
    )
    st.markdown('</div>', unsafe_allow_html=True)

if hist_file is not None and curr_file is not None:
    try:
        with st.spinner("🧠 1. Learning historical vendor-category rules from historical ledger..."):
            hist_wb = openpyxl.load_workbook(io.BytesIO(hist_file.read()))
            learned_rules = extract_historical_rules(hist_wb)
            
        with st.spinner("🔍 2. Auditing current ledger, populating CC vendor names, and styling rows..."):
            curr_wb = openpyxl.load_workbook(io.BytesIO(curr_file.read()))
            out_wb, summary = audit_and_fix_current_ledger(curr_wb, learned_rules)
            
            out_buffer = io.BytesIO()
            out_wb.save(out_buffer)
            out_buffer.seek(0)
            
        st.success(f"🎉 **Audit Complete!** Learned **{len(learned_rules)} vendor rules** from `{hist_file.name}` and audited `{curr_file.name}`.")
        
        # KPI Row
        st.markdown("### 📈 Audit Overview Metrics")
        k1, k2, k3, k4, k5 = st.columns(5)
        with k1:
            st.metric("Total Rows Processed", summary["total_rows"])
        with k2:
            st.metric("CC Payments Fixed", summary["cc_fixed"], delta=f"{summary['cc_fixed']} updated", delta_color="normal")
        with k3:
            st.metric("Category Mismatches", summary["mismatch_count"], delta=f"{summary['mismatch_count']} vs history", delta_color="inverse")
        with k4:
            st.metric("Uncategorized Items", summary["uncat_count"], delta=f"{summary['uncat_count']} need action", delta_color="inverse")
        with k5:
            st.metric("Verified & Matched", summary["matched_count"], delta=f"{summary['matched_count']} clean", delta_color="normal")
            
        st.markdown("---")
        
        # Download Section
        d_col1, d_col2 = st.columns([3, 1])
        with d_col1:
            st.markdown("#### 📥 Download Your Audited & Highlighted Excel File")
            st.caption("Includes the new 'Audit & Discrepancies' summary tab and the 100% color-highlighted ledger tab.")
        with d_col2:
            st.download_button(
                label="⬇️ Download Audited Excel",
                data=out_buffer.getvalue(),
                file_name=f"Audited_{curr_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        st.markdown("---")
        
        # Discrepancies Inspector
        st.markdown("### 🔍 Discrepancies & Flagged Entries Inspector")
        df_discrepancies = pd.DataFrame(summary["discrepancies"])
        
        if not df_discrepancies.empty:
            f_col1, f_col2 = st.columns([2, 2])
            with f_col1:
                status_filter = st.selectbox("Filter by Status", ["All"] + list(df_discrepancies["Status"].unique()))
            with f_col2:
                search_query = st.text_input("Search by Vendor, Description, or Account", "")
                
            filtered_df = df_discrepancies.copy()
            if status_filter != "All":
                filtered_df = filtered_df[filtered_df["Status"] == status_filter]
            if search_query:
                filtered_df = filtered_df[
                    filtered_df["Vendor"].astype(str).str.contains(search_query, case=False, na=False) |
                    filtered_df["Description"].astype(str).str.contains(search_query, case=False, na=False) |
                    filtered_df["Account"].astype(str).str.contains(search_query, case=False, na=False)
                ]
                
            st.dataframe(
                filtered_df[["Row", "Date", "Account", "Type", "Vendor", "Amount", "Status", "Issue", "Action"]].style.format({"Amount": "${:,.2f}"}),
                use_container_width=True,
                height=350
            )
        else:
            st.info("🎉 No discrepancies found! All transactions match your historical rules.")
            
        with st.expander(f"📚 View Dynamically Learned Rules ({len(learned_rules)} Vendors from Historical File)"):
            rules_list = []
            for v_name, v_info in learned_rules.items():
                rules_list.append({
                    "Vendor / Pattern": v_name,
                    "Historical Category": v_info["primary_category"],
                    "Past Occurrences": v_info["occurrences"]
                })
            rules_df = pd.DataFrame(rules_list)
            st.dataframe(rules_df.sort_values(by="Past Occurrences", ascending=False), use_container_width=True, height=250)

    except Exception as e:
        st.error(f"Error executing audit: {str(e)}")
elif hist_file is None or curr_file is None:
    st.info("💡 **Ready to audit**: Please upload both your **Historical Ledger** (Step 1) and your **Current Ledger** (Step 2) above to run the audit.")
