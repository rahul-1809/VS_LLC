import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import os

# Canonical / Historical Vendor-to-Category mapping
HISTORICAL_VENDOR_CATEGORY_MAP = {
    # CC Payments
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
    # Food & Client Meals
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
    # Utilities
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
    # Vendors (AP / COGS)
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
    """Clean and normalize vendor/merchant string from Name or Description."""
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

def process_general_ledger(input_filepath, output_filepath):
    print(f"Loading {input_filepath}...")
    wb = openpyxl.load_workbook(input_filepath)
    sheet = wb["Sheet1"]
    
    header_bg = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_cc_updated = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")      # Cyan / Info
    fill_mismatch = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")        # Yellow / Warning
    fill_uncategorized = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")   # Red / Danger
    fill_matched = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid")         # Light Green / Success
    
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
    
    # Set Audit Headers at Row 5, Col 11 and 12
    cell_k5 = sheet.cell(row=5, column=11, value="Audit Status")
    cell_l5 = sheet.cell(row=5, column=12, value="Audit Details & Suggested Category")
    for cell in [cell_k5, cell_l5]:
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = header_bg
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    discrepancy_list = []
    cc_updates_count = 0
    mismatch_count = 0
    uncat_count = 0
    matched_count = 0
    
    # Scan through 100% of all transaction rows
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
        
        # Skip headers/beginning balances
        if not c_dist or c_dist in ["Beginning Balance", "Distribution account"]:
            continue
        
        str_dist = str(c_dist).strip()
        str_type = str(c_type).strip() if c_type else ""
        str_name = str(c_name).strip() if c_name else ""
        str_desc = str(c_desc).strip() if c_desc else ""
        str_split = str(c_split).strip() if c_split else ""
        amt_val = float(c_amt) if c_amt is not None else 0.0
        
        # 1. CC Payment Detection & Auto-Population
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
            cc_updates_count += 1
            discrepancy_list.append({
                "row": r, "date": c_date, "account": str_dist, "type": str_type,
                "name": target_cc_name, "desc": str_desc, "split": str_split,
                "amt": amt_val, "status": status, "issue": "Missing CC Vendor Name",
                "action": f"Updated Name to '{target_cc_name}'"
            })
        else:
            vendor_key = normalize_vendor(str_name, str_desc)
            
            # Check Uncategorized
            if "Uncategorized" in str_dist or "Uncategorized" in str_split:
                status = "UNCATEGORIZED"
                note = "Uncategorized transaction requires reconciliation / invoice mapping."
                row_fill = fill_uncategorized
                row_font = font_danger
                uncat_count += 1
                discrepancy_list.append({
                    "row": r, "date": c_date, "account": str_dist, "type": str_type,
                    "name": str_name or vendor_key, "desc": str_desc, "split": str_split,
                    "amt": amt_val, "status": status, "issue": "Uncategorized Transaction",
                    "action": "Assign to specific client invoice or expense account"
                })
            elif vendor_key in HISTORICAL_VENDOR_CATEGORY_MAP:
                expected_category = HISTORICAL_VENDOR_CATEGORY_MAP[vendor_key]
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
                    discrepancy_list.append({
                        "row": r, "date": c_date, "account": str_dist, "type": str_type,
                        "name": str_name or vendor_key, "desc": str_desc, "split": str_split,
                        "amt": amt_val, "status": status, "issue": f"Mismatched Category (Found: {curr_category})",
                        "action": f"Reclassify to '{expected_category}'"
                    })
                else:
                    status = "MATCHED"
                    note = f"Matched to historical category: {expected_category}"
                    row_fill = fill_matched
                    row_font = font_success
                    matched_count += 1
            else:
                # All other valid transactions
                status = "MATCHED"
                note = "Standard ledger entry"
                row_fill = fill_matched
                row_font = font_success
                matched_count += 1
        
        # Apply to columns K and L
        cell_k = sheet.cell(row=r, column=11, value=status)
        cell_l = sheet.cell(row=r, column=12, value=note)
        cell_k.font = row_font
        cell_l.font = font_regular
        
        # Color every cell in this row
        for col_idx in range(2, 13):
            sheet.cell(row=r, column=col_idx).fill = row_fill
    
    # CREATE AUDIT SUMMARY TAB
    if "Audit & Discrepancies" in wb.sheetnames:
        del wb["Audit & Discrepancies"]
    audit_ws = wb.create_sheet(title="Audit & Discrepancies", index=0)
    audit_ws.views.sheetView[0].showGridLines = True
    
    # Title
    audit_ws.merge_cells("A1:H1")
    title_cell = audit_ws["A1"]
    title_cell.value = "VIRTUSSOLUTIONS LLC - General Ledger Audit & Compliance Summary"
    title_cell.font = Font(name="Arial", size=14, bold=True, color="1F4E79")
    
    audit_ws["A2"] = "Period: July 2026 | Accounting Basis: Accrual Basis"
    audit_ws["A2"].font = Font(name="Arial", size=10, italic=True, color="595959")
    
    kpis = [
        ("Total Discrepancies / Flags", len(discrepancy_list), "FFF3CD", "856404"),
        ("CC Vendor Payments Fixed", cc_updates_count, "D1ECF1", "0C5460"),
        ("Category Inconsistencies", mismatch_count, "FFF3CD", "856404"),
        ("Uncategorized Inflows/Outflows", uncat_count, "F8D7DA", "721C24"),
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
    
    for row_idx, item in enumerate(discrepancy_list, 8):
        audit_ws.cell(row=row_idx, column=1, value=item["row"]).alignment = Alignment(horizontal="center")
        audit_ws.cell(row=row_idx, column=2, value=item["date"]).alignment = Alignment(horizontal="center")
        audit_ws.cell(row=row_idx, column=3, value=item["account"])
        audit_ws.cell(row=row_idx, column=4, value=item["type"])
        audit_ws.cell(row=row_idx, column=5, value=item["name"]).font = font_bold
        audit_ws.cell(row=row_idx, column=6, value=item["desc"])
        c_amt = audit_ws.cell(row=row_idx, column=7, value=item["amt"])
        c_amt.number_format = "$#,##0.00;($#,##0.00);\"-\""
        
        c_stat = audit_ws.cell(row=row_idx, column=8, value=item["status"])
        if item["status"] == "UNCATEGORIZED":
            c_stat.fill = fill_uncategorized
            c_stat.font = font_danger
        elif item["status"] == "CATEGORY MISMATCH":
            c_stat.fill = fill_mismatch
            c_stat.font = font_warning
        else:
            c_stat.fill = fill_cc_updated
            c_stat.font = font_info
        c_stat.alignment = Alignment(horizontal="center")
        
        audit_ws.cell(row=row_idx, column=9, value=item["issue"])
        audit_ws.cell(row=row_idx, column=10, value=item["action"]).font = font_bold
        
        for c in range(1, 11):
            audit_ws.cell(row=row_idx, column=c).border = thin_border
    
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
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)
    
    wb.save(output_filepath)
    print(f"Audit complete! Saved output to: {output_filepath}")
    print(f" - 100% of rows evaluated: {matched_count + len(discrepancy_list)} rows")
    print(f" - Credit Card Payments Vendor Names Fixed: {cc_updates_count}")
    print(f" - Category Inconsistencies Flagged: {mismatch_count}")
    print(f" - Uncategorized Items Flagged: {uncat_count}")
    print(f" - Matched Transactions: {matched_count}")

if __name__ == "__main__":
    src_file = "VIRTUSSOLUTIONS LLC_General Ledger (1).xlsx"
    out_file = "VIRTUSSOLUTIONS_LLC_Audited_General_Ledger.xlsx"
    process_general_ledger(src_file, out_file)
