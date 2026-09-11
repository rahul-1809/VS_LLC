import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import io
import os
import json
import collections
from groq import Groq

# Page Configuration
st.set_page_config(
    page_title="AI General Ledger Auditor (Groq LLM)",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #F55036;
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
        background-color: #F55036;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-weight: 700;
        font-size: 0.85rem;
        margin-right: 6px;
    }
    .ai-badge {
        background: linear-gradient(135deg, #F55036 0%, #FF8C00 100%);
        color: white;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# SECURE API KEY RESOLUTION
# -----------------------------------------------------------------------------
def get_groq_api_key():
    # 1. Streamlit Secrets
    if "GROQ_API_KEY" in st.secrets:
        return st.secrets["GROQ_API_KEY"]
    # 2. Environment Variable (.env)
    if os.environ.get("GROQ_API_KEY"):
        return os.environ.get("GROQ_API_KEY")
    # 3. Local .env file fallback
    if os.path.exists(".env"):
        with open(".env") as f:
            for line in f:
                if line.startswith("GROQ_API_KEY="):
                    return line.strip().split("=", 1)[1].strip("\"'")
    return ""

# -----------------------------------------------------------------------------
# HISTORICAL KNOWLEDGE EXTRACTION
# -----------------------------------------------------------------------------
def extract_historical_context(hist_wb):
    """Extract summary of vendor -> category associations from historical data."""
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
        vendor_id = c_name if c_name and c_name != "None" else c_desc
        
        if vendor_id and category and "Uncategorized" not in category:
            learned[vendor_id[:40]][category] += 1
            
    summary_map = {}
    for v, cat_counts in learned.items():
        top_cat = cat_counts.most_common(1)[0][0]
        summary_map[v] = {
            "category": top_cat,
            "occurrences": sum(cat_counts.values())
        }
    return summary_map

# -----------------------------------------------------------------------------
# GROQ LLM BATCH AUDIT ENGINE
# -----------------------------------------------------------------------------
def run_groq_llm_audit(client, model_name, transactions_chunk, historical_context):
    """Send a batch of transactions to Groq LLM for intelligent audit & reconciliation."""
    
    prompt = f"""
You are a CPA and Senior Accounting Auditor.
You are given:
1. A summary of HISTORICAL VENDOR CATEGORIES learned from past approved company books.
2. A list of CURRENT LEDGER TRANSACTIONS to audit.

### AUDITING & VERIFICATION RULES:
1. CREDIT CARD PAYMENTS RULE:
   - If a transaction is a credit card payment/transfer (Type 'Credit Card Payment' or split to 'Amex CC' / 'BOA CC' or description indicates paying CC), status must be 'CC VENDOR UPDATED', and recommend setting Vendor to 'Amex CC' or 'BOA CC'.

2. CLIENT REVENUE EQUIVALENCE:
   - 'Accounts Receivable (A/R)' and 'Services' (Revenue) are the two valid sides of the client invoicing & collection cycle. DO NOT flag an invoice or payment as a mismatch between A/R and Services. Both are valid.

3. VENDOR PAYABLES EQUIVALENCE:
   - 'Accounts Payable (A/P)', 'COGS - Vendor', and 'Consulting Fees' are valid counterparts of the vendor subcontractor cycle. DO NOT flag normal bill recordings as mismatches between AP and COGS.

4. SUB-ACCOUNT EQUIVALENCE:
   - Short sub-account names (e.g. 'Phone service' vs 'Utilities:Phone service', 'Other' vs 'Utilities:Other', 'Bank fees & service charges' vs 'General business expenses:Bank fees & service charges') are MATCHED.

5. CATEGORY MISMATCH DETECTION:
   - If a restaurant/food order (e.g. DoorDash, Uber Eats) is placed in 'Entertainment', flag as 'CATEGORY MISMATCH' and recommend 'Meals & Entertainment:Meals with clients'.
   - If software/AI tool (e.g. OpenAI ChatGPT, Anthropic Claude) is in 'Dues & subscriptions', flag as 'CATEGORY MISMATCH' and recommend 'Office expenses:Software & apps'.
   - If maintenance/pest control is in 'Office supplies', flag as 'CATEGORY MISMATCH' and recommend 'Repairs & maintenance'.
   - If 'Uncategorized Income' or 'Uncategorized Expense', status is 'UNCATEGORIZED'.
   - Otherwise, if category aligns with history, status is 'MATCHED'.

HISTORICAL KNOWLEDGE BASE (Sample of learned rules):
{json.dumps(dict(list(historical_context.items())[:60]), indent=1)}

TRANSACTIONS TO AUDIT:
{json.dumps(transactions_chunk, indent=1)}

Respond ONLY with a valid JSON object matching this schema:
{{
  "results": [
    {{
      "row_id": <int>,
      "status": "MATCHED" | "CC VENDOR UPDATED" | "CATEGORY MISMATCH" | "UNCATEGORIZED",
      "updated_vendor_name": "<string if CC payment or blank>",
      "issue": "<concise description of issue if any or None>",
      "recommended_action": "<exact action / suggested category>"
    }}
  ]
}}
"""
    try:
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": "You are an expert CPA and automated financial audit agent. Always respond with precise, valid JSON."},
                {"role": "user", "content": prompt}
            ],
            model=model_name,
            response_format={"type": "json_object"},
            temperature=0.1
        )
        content = response.choices[0].message.content
        data = json.loads(content)
        return data.get("results", [])
    except Exception as e:
        st.warning(f"LLM API batch fallback triggered: {str(e)}")
        # Graceful fallback heuristic if LLM times out
        fallback_results = []
        for tx in transactions_chunk:
            r_id = tx["row_id"]
            d_type = tx.get("type", "")
            d_split = tx.get("split", "")
            d_desc = tx.get("desc", "")
            if "Credit Card" in d_type or "Amex CC" in d_split or "BOA CC" in d_split:
                cc = "Amex CC" if "Amex" in d_split or "AMEX" in d_desc else "BOA CC"
                fallback_results.append({
                    "row_id": r_id, "status": "CC VENDOR UPDATED", "updated_vendor_name": cc,
                    "issue": "Missing CC Vendor Name", "recommended_action": f"Set Vendor Name to '{cc}'"
                })
            elif "Uncategorized" in d_split or "Uncategorized" in tx.get("account", ""):
                fallback_results.append({
                    "row_id": r_id, "status": "UNCATEGORIZED", "updated_vendor_name": "",
                    "issue": "Uncategorized Entry", "recommended_action": "Reconcile invoice"
                })
            else:
                fallback_results.append({
                    "row_id": r_id, "status": "MATCHED", "updated_vendor_name": "",
                    "issue": None, "recommended_action": "Verified"
                })
        return fallback_results


# -----------------------------------------------------------------------------
# WORKBOOK GENERATION & STYLING
# -----------------------------------------------------------------------------
def build_audited_workbook(curr_wb, audit_map):
    sheet = curr_wb.active
    
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
    
    cell_k5 = sheet.cell(row=5, column=11, value="AI Audit Status")
    cell_l5 = sheet.cell(row=5, column=12, value="AI Auditor Findings & Recommended Category")
    for cell in [cell_k5, cell_l5]:
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = header_bg
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    discrepancy_list = []
    cc_fixed = 0
    mismatch_count = 0
    uncat_count = 0
    matched_count = 0
    
    for r in range(6, sheet.max_row + 1):
        c_dist = sheet.cell(row=r, column=2).value
        if not c_dist or c_dist in ["Beginning Balance", "Distribution account"]:
            continue
            
        c_date = sheet.cell(row=r, column=3).value
        c_type = sheet.cell(row=r, column=4).value
        cell_name = sheet.cell(row=r, column=6)
        c_desc = sheet.cell(row=r, column=7).value
        c_split = sheet.cell(row=r, column=8).value
        c_amt = sheet.cell(row=r, column=9).value
        amt_val = float(c_amt) if c_amt is not None else 0.0
        
        res = audit_map.get(r, {"status": "MATCHED", "issue": None, "recommended_action": "Verified", "updated_vendor_name": ""})
        status = res.get("status", "MATCHED")
        action = res.get("recommended_action", "Verified")
        issue = res.get("issue", "")
        updated_name = res.get("updated_vendor_name", "")
        
        if status == "CC VENDOR UPDATED":
            if updated_name:
                cell_name.value = updated_name
            row_fill = fill_cc_updated
            row_font = font_info
            cc_fixed += 1
            discrepancy_list.append({
                "Row": r, "Date": c_date, "Account": c_dist, "Type": c_type,
                "Vendor": cell_name.value or updated_name, "Description": c_desc,
                "Split": c_split, "Amount": amt_val, "Status": status,
                "Issue": issue or "Missing CC Vendor Name", "Action": action
            })
        elif status == "UNCATEGORIZED":
            row_fill = fill_uncategorized
            row_font = font_danger
            uncat_count += 1
            discrepancy_list.append({
                "Row": r, "Date": c_date, "Account": c_dist, "Type": c_type,
                "Vendor": cell_name.value or "", "Description": c_desc,
                "Split": c_split, "Amount": amt_val, "Status": status,
                "Issue": issue or "Uncategorized Entry", "Action": action
            })
        elif status == "CATEGORY MISMATCH":
            row_fill = fill_mismatch
            row_font = font_warning
            mismatch_count += 1
            discrepancy_list.append({
                "Row": r, "Date": c_date, "Account": c_dist, "Type": c_type,
                "Vendor": cell_name.value or "", "Description": c_desc,
                "Split": c_split, "Amount": amt_val, "Status": status,
                "Issue": issue or "Category Mismatch", "Action": action
            })
        else:
            row_fill = fill_matched
            row_font = font_success
            matched_count += 1
            
        cell_k = sheet.cell(row=r, column=11, value=status)
        cell_l = sheet.cell(row=r, column=12, value=action)
        cell_k.font = row_font
        cell_l.font = font_regular
        
        for col_idx in range(2, 13):
            sheet.cell(row=r, column=col_idx).fill = row_fill
            
    if "Audit & Discrepancies" in curr_wb.sheetnames:
        del curr_wb["Audit & Discrepancies"]
    audit_ws = curr_wb.create_sheet(title="Audit & Discrepancies", index=0)
    audit_ws.views.sheetView[0].showGridLines = True
    
    audit_ws["A1"] = "Groq LLM General Ledger AI Audit & Compliance Summary"
    audit_ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F4E79")
    audit_ws["A2"] = f"Total Transactions Audited: {len(discrepancy_list) + matched_count} | Issues / Updates: {len(discrepancy_list)}"
    audit_ws["A2"].font = Font(name="Arial", size=10, italic=True, color="595959")
    
    kpis = [
        ("Total Flags / Updates", len(discrepancy_list), "FFF3CD", "856404"),
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
        
    headers = ["Row #", "Date", "Account", "Type", "Vendor / Name", "Description", "Amount", "Status", "Identified Issue", "AI Recommendation"]
    for col_idx, h_text in enumerate(headers, 1):
        cell = audit_ws.cell(row=7, column=col_idx, value=h_text)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = header_bg
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for idx, item in enumerate(discrepancy_list, 8):
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
        "total_rows": len(discrepancy_list) + matched_count,
        "cc_fixed": cc_fixed,
        "mismatch_count": mismatch_count,
        "uncat_count": uncat_count,
        "matched_count": matched_count,
        "discrepancies": discrepancy_list
    }
    return curr_wb, summary_data


# -----------------------------------------------------------------------------
# STREAMLIT UI
# -----------------------------------------------------------------------------
st.markdown('<div class="main-header">⚡ AI General Ledger Auditor <span class="ai-badge">Powered by Groq LLM</span></div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload your <b>Historical Ledger</b> and <b>Current Ledger</b>. The Groq LLM dynamically understands transaction context, enforces credit card vendor naming, verifies category classifications, and exports a 100% color-audited workbook.</div>', unsafe_allow_html=True)

# Sidebar
st.sidebar.header("⚡ Groq LLM Configuration")
default_key = get_groq_api_key()
api_key_input = st.sidebar.text_input(
    "Groq API Key (Secure)",
    value=default_key,
    type="password",
    help="Your Groq API key is kept secure and never exposed."
)

model_choice = st.sidebar.selectbox(
    "Groq AI Model",
    ["openai/gpt-oss-120b", "qwen/qwen3.8-27b", "openai/gpt-oss-20b", "qwen/qwen3.6-27b"],
    help="Select the Groq high-speed LLM model to power the audit."
)

st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Audit Color Legend")
st.sidebar.markdown("""
- 🟦 **Blue**: Credit Card Payment (Vendor Name Populated)
- 🟨 **Yellow**: Category Mismatch vs Past History
- 🟥 **Red**: Uncategorized Income / Expense
- 🟩 **Green**: Verified & Matched Entry
""")

# Dual Upload Section
col_hist, col_curr = st.columns(2)

with col_hist:
    st.markdown('<div class="upload-card"><span class="step-badge">STEP 1</span><b>Upload Historical Ledger (Ground Truth)</b>', unsafe_allow_html=True)
    hist_file = st.file_uploader(
        "Upload Past Correctly Classified Ledger (.xlsx or .xls)",
        type=["xlsx", "xls"],
        key="groq_hist_uploader",
        help="Upload past general ledger containing historical category truth."
    )
    st.markdown('</div>', unsafe_allow_html=True)

with col_curr:
    st.markdown('<div class="upload-card"><span class="step-badge">STEP 2</span><b>Upload Current Ledger (To Audit & Fix)</b>', unsafe_allow_html=True)
    curr_file = st.file_uploader(
        "Upload Current Month Ledger (.xlsx or .xls)",
        type=["xlsx", "xls"],
        key="groq_curr_uploader",
        help="Upload the new general ledger file to be audited by Groq LLM."
    )
    st.markdown('</div>', unsafe_allow_html=True)

# Processing
if hist_file is not None and curr_file is not None:
    if not api_key_input:
        st.error("⚠️ Please provide a valid Groq API Key in the sidebar to run the AI audit.")
    else:
        try:
            client = Groq(api_key=api_key_input)
            
            with st.spinner("🧠 1. Learning historical categorization rules from Historical Ledger..."):
                hist_wb = openpyxl.load_workbook(io.BytesIO(hist_file.read()))
                hist_context = extract_historical_context(hist_wb)
                
            with st.spinner(f"⚡ 2. Processing current ledger with Groq LLM ({model_choice})..."):
                curr_wb = openpyxl.load_workbook(io.BytesIO(curr_file.read()))
                sheet = curr_wb.active
                curr_rows = list(sheet.iter_rows(values_only=True))
                
                # Extract transactions
                tx_list = []
                for r_idx, r in enumerate(curr_rows):
                    c_dist = str(r[1]).strip() if r[1] is not None else ""
                    if not c_dist or c_dist in ["Beginning Balance", "Distribution account"]:
                        continue
                    tx_list.append({
                        "row_id": r_idx + 1,
                        "account": c_dist,
                        "date": str(r[2]),
                        "type": str(r[3]),
                        "vendor": str(r[5]).strip() if r[5] else "",
                        "desc": str(r[6]).strip() if r[6] else "",
                        "split": str(r[7]).strip() if r[7] else "",
                        "amount": r[8]
                    })
                
                # Batch processing with progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()
                audit_map = {}
                batch_size = 25
                total_batches = (len(tx_list) + batch_size - 1) // batch_size
                
                for b_idx in range(total_batches):
                    batch = tx_list[b_idx * batch_size : (b_idx + 1) * batch_size]
                    status_text.text(f"⚡ Auditing transactions {b_idx * batch_size + 1} to {min((b_idx + 1) * batch_size, len(tx_list))} of {len(tx_list)}...")
                    batch_results = run_groq_llm_audit(client, model_choice, batch, hist_context)
                    for res in batch_results:
                        audit_map[res["row_id"]] = res
                    progress_bar.progress((b_idx + 1) / total_batches)
                    
                status_text.empty()
                progress_bar.empty()
                
            with st.spinner("🎨 3. Generating color-highlighted Excel workbook & dashboard..."):
                out_wb, summary = build_audited_workbook(curr_wb, audit_map)
                out_buffer = io.BytesIO()
                out_wb.save(out_buffer)
                out_buffer.seek(0)
                
            st.success(f"🎉 **AI Audit Complete!** Evaluated **{summary['total_rows']} transactions** with **Groq LLM ({model_choice})**.")
            
            # KPI Metrics
            st.markdown("### 📈 AI Audit Summary Overview")
            k1, k2, k3, k4, k5 = st.columns(5)
            with k1:
                st.metric("Total Rows Evaluated", summary["total_rows"])
            with k2:
                st.metric("CC Payments Fixed", summary["cc_fixed"], delta=f"{summary['cc_fixed']} updated", delta_color="normal")
            with k3:
                st.metric("Category Inconsistencies", summary["mismatch_count"], delta=f"{summary['mismatch_count']} to review", delta_color="inverse")
            with k4:
                st.metric("Uncategorized Items", summary["uncat_count"], delta=f"{summary['uncat_count']} need action", delta_color="inverse")
            with k5:
                st.metric("Verified & Matched", summary["matched_count"], delta=f"{summary['matched_count']} clean", delta_color="normal")
                
            st.markdown("---")
            
            # Download
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.markdown("#### 📥 Download AI-Audited Excel Workbook")
                st.caption("Contains the 'Audit & Discrepancies' tab and the 100% color-highlighted ledger tab.")
            with d_col2:
                st.download_button(
                    label="⬇️ Download Audited Excel",
                    data=out_buffer.getvalue(),
                    file_name=f"Groq_Audited_{curr_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            st.markdown("---")
            
            # Discrepancies Inspector
            st.markdown("### 🔍 AI Discrepancies & Flagged Entries Inspector")
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
                st.info("🎉 No discrepancies found! All transactions match historical patterns.")

        except Exception as e:
            st.error(f"Error during AI audit: {str(e)}")
elif hist_file is None or curr_file is None:
    st.info("💡 **Ready to audit with Groq AI**: Please upload both your **Historical Ledger** (Step 1) and your **Current Ledger** (Step 2) above to begin.")
