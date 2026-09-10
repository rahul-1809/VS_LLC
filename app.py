import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import io
import re

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
    .upload-box {
        border: 2px dashed #1F4E79;
        border-radius: 10px;
        padding: 2.5rem;
        text-align: center;
        background-color: #F8FAFC;
        margin-bottom: 1.5rem;
    }
    .feature-card {
        background: #ffffff;
        border-radius: 8px;
        padding: 1.2rem;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        height: 100%;
    }
    .feature-title {
        font-weight: 700;
        color: #1F4E79;
        font-size: 1.05rem;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Canonical Vendor-to-Category Mapping Base
DEFAULT_VENDOR_CATEGORY_MAP = {
    # Credit Card Accounts
    "Amex CC": "Amex CC",
    "BOA CC": "BOA CC",
    # Software & SaaS
    "OpenAI": "Office expenses:Software & apps",
    "Anthropic": "Office expenses:Software & apps",
    "Google Workspace": "Office expenses:Software & apps",
    "Microsoft": "Office expenses:Software & apps",
    "LinkedIn": "Office expenses:Software & apps",
    "Intuit QuickBooks": "Dues & subscriptions",
    "HP Instant Ink": "Office expenses:Office supplies",
    "Amazon": "Office expenses:Office supplies",
    # Client Meals & Catering
    "DoorDash": "Meals & Entertainment:Meals with clients",
    "Uber Eats": "Meals & Entertainment:Meals with clients",
    "Club 28 Badminton": "Meals & Entertainment:Meals with clients",
    # Travel & Vehicles
    "Uber": "Travel:Taxis or shared rides",
    "Wawa": "Travel:Auto Expenses",
    "Chevron": "Gas And Fuel",
    "Union 76": "Gas And Fuel",
    "Quik Stop": "Gas And Fuel",
    "Fastrak": "Vehicle expenses:Parking & tolls",
    "SpotHero": "Vehicle expenses:Parking & tolls",
    # Utilities & Telecom
    "Delmarva Power": "Utilities:Electricity",
    "Comcast / Xfinity": "Utilities:Other",
    "AT&T": "Utilities:Phone service",
    "Verizon": "Utilities:Phone service",
    "RingCentral": "Utilities:Phone service",
    "Vonage": "Utilities:Phone service",
    # Payroll & Tax
    "ADP Payroll Processing": "Payroll expenses:Payroll Processing Fee",
    "ADP Payroll Tax": "Payroll wages and tax to pay:Payroll tax to pay",
    "ADP 401k": "401K Payable",
    # Professional, Legal & Insurance
    "CSAA Insurance": "Insurance",
    "Global Immigration Partners": "Immigration Expenses",
    "Banner Pest Services": "Repairs & maintenance",
    # Personal / Draws
    "Synergy Badminton": "Owners Distribution",
    "Freewheel Brewing": "Owners Distribution",
    "Activate Plano": "Shareholders' equity:Distributions",
    "Netflix": "General business expenses:Memberships & subscriptions",
    "Prime Video": "General business expenses:Memberships & subscriptions",
    # Vendors (COGS / AP)
    "Astramind Solutions": "Accounts Payable (A/P)",
    "Aneru LLC": "Accounts Payable (A/P)",
    "PTL Consultancy": "Accounts Payable (A/P)",
    "Sysconsult LLC": "Consulting Fees",
    "Sage Solutions": "Consulting Fees",
    # Clients (AR / Services)
    "American Technology Services LLC": "Accounts Receivable (A/R)",
    "Agilisium Consulting LLC": "Accounts Receivable (A/R)",
    "Medix Staffing Solutions": "Accounts Receivable (A/R)",
    "Nityo Infotech Corporation": "Accounts Receivable (A/R)",
    "Infinite Computer Solutions": "Accounts Receivable (A/R)",
    "Phoenix Staff Inc": "Accounts Receivable (A/R)",
    "Bigbyte": "Accounts Receivable (A/R)",
    "SANA Software Solutions": "Accounts Receivable (A/R)",
    "Vivid Technologies Inc": "Accounts Receivable (A/R)",
    "Vendorpass Inc": "Accounts Receivable (A/R)",
    "Randstad Digital": "Accounts Receivable (A/R)",
    "Sanvi Consulting Services LLC": "Accounts Receivable (A/R)",
}

def normalize_vendor(name, desc):
    """Clean and normalize merchant/vendor string from Name or Description."""
    text = str(name).strip() if name and str(name).strip() and str(name).strip() != "None" else (str(desc).strip() if desc else "")
    if not text:
        return "Unknown"
    
    t_upper = text.upper()
    
    if "AMERICAN EXPRESS" in t_upper or "AMEX" in t_upper:
        return "Amex CC"
    if "BANK OF AMERICA - CREDIT CARD" in t_upper or "BOA CC" in t_upper or "B OF A CREDIT CARD" in t_upper:
        return "BOA CC"
    if "ADP PAYROLL FEES" in t_upper or "ADP FEES" in t_upper:
        return "ADP Payroll Processing"
    if "ADP TAX" in t_upper:
        return "ADP Payroll Tax"
    if "ADP 401K" in t_upper:
        return "ADP 401k"
    if "DOORDASH" in t_upper or "DD *" in t_upper:
        return "DoorDash"
    if "UBER EATS" in t_upper:
        return "Uber Eats"
    if "UBER" in t_upper:
        return "Uber"
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
    if "AMAZON PRIME" in t_upper or "AMAZON DIGI" in t_upper or "AMAZON" in t_upper:
        return "Amazon"
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
    if "INTUIT" in t_upper or "QUICKBOOKS" in t_upper:
        return "Intuit QuickBooks"
    if "NETFLIX" in t_upper:
        return "Netflix"
    if "PRIME VIDEO" in t_upper:
        return "Prime Video"
    if "HP INSTANT INK" in t_upper:
        return "HP Instant Ink"
    
    clean = re.sub(r"^(AplPay|TST\*\s*|DD\s*\*\s*|PY\s*\*\s*|PAYPAL\s*\*|SQ\s*\*)", "", text, flags=re.IGNORECASE).strip()
    return clean[:30]

def audit_and_fix_workbook(wb, vendor_rules):
    sheet = wb.active
    
    # Styling definitions
    header_bg = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_cc_updated = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")
    fill_mismatch = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    fill_uncategorized = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    fill_matched = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid")
    
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
    
    # Audit column headers at row 5
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
    
    # Iterate across ALL rows
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
        
        # Skip headers/beginning balance
        if not c_dist or c_dist in ["Beginning Balance", "Distribution account"]:
            continue
            
        str_dist = str(c_dist).strip()
        str_type = str(c_type).strip() if c_type else ""
        str_name = str(c_name).strip() if c_name else ""
        str_desc = str(c_desc).strip() if c_desc else ""
        str_split = str(c_split).strip() if c_split else ""
        amt_val = float(c_amt) if c_amt is not None else 0.0
        
        # 1. Check Credit Card Payments
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
            note = f"Vendor Name set to '{target_cc_name}'"
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
            vendor_key = normalize_vendor(str_name, str_desc)
            
            # Check Uncategorized
            if "Uncategorized" in str_dist or "Uncategorized" in str_split:
                status = "UNCATEGORIZED"
                note = "Requires invoice or category reconciliation."
                row_fill = fill_uncategorized
                row_font = font_danger
                uncat_count += 1
                audit_results.append({
                    "Row": r, "Date": c_date, "Account": str_dist, "Type": str_type,
                    "Vendor": str_name or vendor_key, "Description": str_desc, "Split": str_split,
                    "Amount": amt_val, "Status": status, "Issue": "Uncategorized Inflow/Outflow",
                    "Action": "Reconcile with specific customer invoice or account"
                })
            elif vendor_key in vendor_rules:
                expected_category = vendor_rules[vendor_key]
                curr_category = str_split if str_dist in ["Bank of America", "Amex CC", "BOA CC"] else str_dist
                
                is_match = False
                if expected_category.lower() in curr_category.lower() or curr_category.lower() in expected_category.lower():
                    is_match = True
                if curr_category in ["Accounts Receivable (A/R)", "Accounts Payable (A/P)"]:
                    is_match = True
                
                if not is_match:
                    status = "CATEGORY MISMATCH"
                    note = f"Expected '{expected_category}', found '{curr_category}'."
                    row_fill = fill_mismatch
                    row_font = font_warning
                    mismatch_count += 1
                    audit_results.append({
                        "Row": r, "Date": c_date, "Account": str_dist, "Type": str_type,
                        "Vendor": str_name or vendor_key, "Description": str_desc, "Split": str_split,
                        "Amount": amt_val, "Status": status, "Issue": f"Mismatched Category (Found: {curr_category})",
                        "Action": f"Reclassify to '{expected_category}'"
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
        
        # Apply Status to columns K and L
        cell_k = sheet.cell(row=r, column=11, value=status)
        cell_l = sheet.cell(row=r, column=12, value=note)
        cell_k.font = row_font
        cell_l.font = font_regular
        
        # Color all cells across row
        for col_idx in range(2, 13):
            sheet.cell(row=r, column=col_idx).fill = row_fill
            
    # Add or update Audit Tab
    if "Audit & Discrepancies" in wb.sheetnames:
        del wb["Audit & Discrepancies"]
    audit_ws = wb.create_sheet(title="Audit & Discrepancies", index=0)
    audit_ws.views.sheetView[0].showGridLines = True
    
    # Title
    audit_ws["A1"] = "General Ledger Audit & Compliance Summary"
    audit_ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F4E79")
    audit_ws["A2"] = f"Total Transactions Processed: {len(audit_results) + matched_count} | Issues/Updates: {len(audit_results)}"
    audit_ws["A2"].font = Font(name="Arial", size=10, italic=True, color="595959")
    
    # KPI headers
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
            
    # Auto-adjust column widths
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
    
    return wb, summary_data


# -----------------------------------------------------------------------------
# STREAMLIT UI (CLEAN LANDING PAGE BY DEFAULT)
# -----------------------------------------------------------------------------
st.markdown('<div class="main-header">📊 General Ledger Auditor & Auto-Fixer</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload your monthly QuickBooks or General Ledger Excel file to automatically audit categorizations, fix missing credit card vendor names, and download a color-coded compliance report.</div>', unsafe_allow_html=True)

# Sidebar
st.sidebar.header("⚙️ Controls")
st.sidebar.markdown("---")

st.sidebar.subheader("🎨 Audit Color Legend")
st.sidebar.markdown("""
- 🟦 **Blue**: Credit Card Payment (Vendor Name Populated)
- 🟨 **Yellow**: Category Mismatch vs History
- 🟥 **Red**: Uncategorized Income / Expense
- 🟩 **Green**: Verified & Matched Entry
""")

uploaded_file = st.file_uploader(
    "📁 Upload General Ledger Excel File (.xlsx or .xls)",
    type=["xlsx", "xls"],
    help="Select and upload any General Ledger workbook to begin the automated audit."
)

if uploaded_file is not None:
    try:
        with st.spinner("🔍 Auditing 100% of rows and verifying vendor categories..."):
            file_bytes = uploaded_file.read()
            file_name = uploaded_file.name
            
            in_wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
            out_wb, summary = audit_and_fix_workbook(in_wb, DEFAULT_VENDOR_CATEGORY_MAP)
            
            out_buffer = io.BytesIO()
            out_wb.save(out_buffer)
            out_buffer.seek(0)
            
        st.success(f"✅ Audit completed successfully for: **{file_name}**")
        
        # KPI Row
        st.markdown("### 📈 Audit Summary Overview")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("Total Rows Evaluated", summary["total_rows"])
        with col2:
            st.metric("CC Payments Fixed", summary["cc_fixed"], delta=f"{summary['cc_fixed']} updated", delta_color="normal")
        with col3:
            st.metric("Category Mismatches", summary["mismatch_count"], delta=f"{summary['mismatch_count']} to review", delta_color="inverse")
        with col4:
            st.metric("Uncategorized Items", summary["uncat_count"], delta=f"{summary['uncat_count']} need action", delta_color="inverse")
        with col5:
            st.metric("Verified & Matched", summary["matched_count"], delta=f"{summary['matched_count']} clean", delta_color="normal")
            
        st.markdown("---")
        
        # Download Section
        d_col1, d_col2 = st.columns([3, 1])
        with d_col1:
            st.markdown("#### 📥 Download Your Highlighted & Audited Excel File")
            st.caption("Contains the new 'Audit & Discrepancies' tab and color-coded General Ledger tab.")
        with d_col2:
            st.download_button(
                label="⬇️ Download Audited Excel",
                data=out_buffer.getvalue(),
                file_name=f"Audited_{file_name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        st.markdown("---")
        
        # Interactive Discrepancies Table
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
            st.info("🎉 No discrepancies found! All entries match historical rules.")
            
        with st.expander("📚 View Historical Vendor $\\rightarrow$ Category Mapping Rules"):
            rules_df = pd.DataFrame(list(DEFAULT_VENDOR_CATEGORY_MAP.items()), columns=["Vendor / Merchant", "Expected Category / Account"])
            st.dataframe(rules_df, use_container_width=True, height=250)

    except Exception as e:
        st.error(f"Error processing workbook: {str(e)}")
else:
    # Fresh Landing Page Explanation
    st.info("👆 Please upload a General Ledger `.xlsx` or `.xls` file above to start.")
    
    st.markdown("### 🚀 What this tool does automatically:")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class="feature-card">
            <div class="feature-title">💳 1. Credit Card Vendor Names</div>
            <p>Automatically detects credit card payments and populates the missing Vendor Name with the Credit Card name (e.g. <code>Amex CC</code> or <code>BOA CC</code>).</p>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="feature-card">
            <div class="feature-title">🔄 2. Historical Category Matching</div>
            <p>Cross-references vendors against historical categories and flags inconsistencies (e.g. DoorDash under Entertainment vs Meals with clients).</p>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class="feature-card">
            <div class="feature-title">🎨 3. 100% Color Highlighting</div>
            <p>Generates an audited workbook with an executive Discrepancies dashboard, row-level color coding, and direct export to Excel.</p>
        </div>
        """, unsafe_allow_html=True)
