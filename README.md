# General Ledger Auditor & Auto-Fixer 📊

An automated financial compliance tool and Streamlit web application for QuickBooks / General Ledger exports.

## Key Features

1. **Credit Card Payment Vendor Population**:
   - Automatically detects payments made to Credit Cards (`Amex CC`, `BOA CC`, etc.) and populates the missing Vendor Name with the exact Credit Card name.
2. **Historical Vendor-to-Category Reconciliation**:
   - Matches every transaction against historical categorization rules.
   - Detects category discrepancies (e.g. DoorDash under Entertainment vs Client Meals, ChatGPT under Subscriptions vs Software & Apps, Pest Control under Supplies vs Maintenance).
3. **100% Visual Row Highlighting & Status**:
   - 🟦 **Light Blue (`#D1ECF1`)**: Credit Card Payment with Vendor Name Auto-Populated.
   - 🟨 **Light Yellow (`#FFF3CD`)**: Category Mismatch vs Past History.
   - 🟥 **Light Red (`#F8D7DA`)**: Uncategorized Income / Expense requiring review.
   - 🟩 **Light Green (`#E8F8F5`)**: Verified & Matched Entries.
   - Adds Column K (`Audit Status`) and Column L (`Audit Details & Suggested Category`).
4. **Executive Dashboard & Discrepancies Tab**:
   - Generates an automated KPI dashboard and detailed issues table in the exported Excel workbook.
5. **Interactive Streamlit Web App**:
   - Upload any `.xlsx` or `.xls` ledger file.
   - Live KPI metric cards and searchable/filterable table.
   - 1-click download of the audited workbook.

---

## Installation & Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/rahul-1809/VS_LLC.git
   cd VS_LLC
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

### Option 1: Run the Streamlit Web Application
```bash
streamlit run app.py
```
Open your browser at `http://localhost:8501`.

### Option 2: Run the CLI Script Directly
```bash
python audit_and_fix_general_ledger.py
```
Generates `VIRTUSSOLUTIONS_LLC_Audited_General_Ledger.xlsx`.

---

## Project Structure
```
├── app.py                                          # Streamlit web application
├── audit_and_fix_general_ledger.py                 # Standalone audit CLI script
├── requirements.txt                                # Python dependencies
├── .gitignore                                      # Ignored build and temporary files
├── README.md                                       # Project documentation
├── VIRTUSSOLUTIONS LLC_General Ledger (1).xlsx    # Sample ledger input
└── VIRTUSSOLUTIONS_LLC_Audited_General_Ledger.xlsx # Audited output workbook
```
