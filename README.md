# General Ledger Auditor & Auto-Fixer 📊

An automated financial compliance tool and Streamlit web applications for QuickBooks / General Ledger exports.

---

## 🌟 Available Applications

### 1. **Groq LLM Powered AI Auditor** (`app_groq.py`) ⚡
Uses **Groq LLM (`openai/gpt-oss-120b`, `qwen/qwen3.8-27b`)** to dynamically reason over transactions, compare against historical truth data, and audit the ledger.

**Run the Groq AI App**:
```bash
streamlit run app_groq.py
```

### 2. **Rule-Based Fast Auditor** (`app.py`) 🚀
Deterministic, high-speed engine that parses historical ledgers and audits current ledgers instantly without external API calls.

**Run the Standard App**:
```bash
streamlit run app.py
```

### 3. **CLI Terminal Tool** (`audit_and_fix_general_ledger.py`) 💻
```bash
python audit_and_fix_general_ledger.py
```

---

## 🔒 API Key Security
- The Groq API key is kept completely secure and is read automatically from `.env` / `.streamlit/secrets.toml`.
- Secret files and keys are **100% ignored in `.gitignore`** and never committed to GitHub.

---

## 🎨 Color-Coded Audit System (100% Row Coverage)
- 🟦 **Light Blue (`#D1ECF1`)**: Credit Card Payment with Vendor Name Auto-Populated (`Amex CC` / `BOA CC`).
- 🟨 **Light Yellow (`#FFF3CD`)**: Category Inconsistency vs Past History.
- 🟥 **Light Red (`#F8D7DA`)**: Uncategorized Income / Expense.
- 🟩 **Light Green (`#E8F8F5`)**: 100% Verified & Matched Entry.
- **Columns K & L**: Automatically appended with `Audit Status` and `Audit Details & Suggested Category`.

---

## 🚀 Setup & Installation
```bash
git clone https://github.com/rahul-1809/VS_LLC.git
cd VS_LLC
pip install -r requirements.txt
```
