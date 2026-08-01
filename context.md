# Bank Statement → Excel Converter — Project Context

> Share this file with a new Claude session when tokens run out.
> Say: "Here's the context for the project we were building. Continue from here."

---

## What This Project Is

A **Streamlit app** that converts bank statement PDFs into styled Excel files.
Built by Hexi (final-year MCA student, intern at Bizmetric Mumbai) as a tool for his uncle to use.

**Core idea:** Uncle uploads a bank PDF → app extracts account details + all transactions → downloads as Excel with two sheets.

---

## Current File: `streamlit_app.py`

Single-file Streamlit app. No separate backend — everything is in one file.

### How to run
```bash
pip install streamlit pdfplumber openpyxl pandas pdf2image pytesseract pdfminer.six
# Tesseract OCR must also be installed on the system:
# Linux: sudo apt install tesseract-ocr
# Mac:   brew install tesseract
streamlit run streamlit_app.py
```

---

## App Flow

### Homepage (before upload)
1. **Mode selector** (radio, horizontal):
   - `📊 Parse & Download Excel` — deterministic parser, no AI, outputs Excel
   - `🤖 Extract Text for AI` — OCRs/extracts raw text, builds a prompt for ChatGPT/Claude/Gemini
2. **Bank selector** (selectbox): UCO Bank | Canara Bank | SBI (Coming Soon) | HDFC (Coming Soon)
3. **File uploader** — PDF only
4. **PDF password field** — optional, for password-protected PDFs
5. **"🚀 Process PDF" submit button** — nothing runs until this is clicked

### Mode A — Parse & Download Excel
- Stat cards: Transactions count / Total Withdrawals (red) / Total Deposits (green) / Closing Balance
- **Download Excel** button (primary, top)
- Two tabs:
  - `🏛️ Account Details` — key-value grid using `st.columns([1, 2])`
  - `💳 Transactions (N)` — `st.dataframe` with 520px height, frozen header

### Mode B — Extract Text for AI
- Stat cards: Pages / Characters / Prompt Size
- Info box explaining copy-paste workflow
- Two download buttons: "Download Full Prompt (.txt)" and "Download Raw Text Only"
- Two preview tabs: Prompt Preview / Raw Text Preview (first 3000 chars each)

---

## Excel Output (2 sheets)

**Sheet 1 — Account Details**
- Title row merged A1:B1 — `"{Bank Name} — Account Statement Details"`
- Rows: Field | Value, alternating `EBF8FF` fill, thin border

**Sheet 2 — Transactions**
- Columns: Date | Particulars | Withdrawals (₹) | Deposits (₹) | Balance (₹)
- Header: navy `1E3A5F` fill, white bold font
- Withdrawals: red `C53030` font, right-aligned
- Deposits: green `276749` font, right-aligned
- Balance: bold, right-aligned
- Alternating row fill `F0F4F8`
- Freeze panes at A2
- Column widths: [15, 55, 18, 18, 18]

---

## Architecture — Router Pattern

```python
def parse_bank_statement(bank_name, file_bytes, page1_text, password):
    if bank_name == "UCO Bank":
        account_info = extract_uco_account_info(page1_text)
        transactions = extract_uco_transactions(file_bytes, password)
        return account_info, transactions

    elif bank_name == "Canara Bank":
        account_info = extract_canara_info_from_bytes(file_bytes)
        transactions = extract_canara_transactions(file_bytes, password)
        return account_info, transactions

    elif bank_name == "SBI (Coming Soon)":
        st.warning("SBI parsing not yet implemented.")
        st.stop()
```

To add a new bank: write `extract_{bank}_account_info()` and `extract_{bank}_transactions()`, then add an `elif` branch here and add to the selectbox options list.

---

## Bank Parsers

### UCO Bank — Deterministic (pdfplumber)
- PDF type: **text-based** (selectable text, no OCR needed)
- Uses `pdfplumber.extract_words()` grouped by Y-position
- Date format: `DD-Mon-YYYY` (e.g. `02-Apr-2025`)
- Column X-boundaries (pdfplumber coords on this PDF):
  - Date: x0 < 85
  - Particulars: 85 ≤ x0 < 310 (and x1 ≤ 360)
  - Withdrawal: x1 < 440 and x1 > 360
  - Deposit: x1 ≥ 440 and x1 < 1440
  - Balance: x1 > 540

**Key logic in `parse_transaction_line()`:**
```python
for w in line_words[1:]:
    is_numeric = NUMBER_RE.match(w['text'].replace(',', ''))
    if w['x1'] > 360 and is_numeric:
        numbers.append((w['x1'], w['text']))
    else:
        particulars_words.append(w['text'])

# bucket numbers by x1:
# x1 < 440 → withdrawal
# x1 >= 440 → deposit
# last number always → balance
```

**Account info extraction** (`extract_uco_account_info`):
- `page1_text` from `pdfplumber` page 1
- Two-column layout: name appears ABOVE the "Name\n" label in OCR order
- Name regex: `r'((?:[A-Z][A-Z\s&\./]+\n)+)Name\n'` — grabs ALL-CAPS lines before label
- Address: finds second `Address\n` occurrence, grabs lines until IFSC/MICR/A/c
- Email: `r'E-Mail ID\s+([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})'` — stops at first match

**Continuation lines:** If a line has no date but has text in x0 < 360, it's appended to previous transaction's particulars.

---

### Canara Bank — OCR-based (pdf2image + pytesseract)
- PDF type: **image-based** (scanned, zero chars extractable by pdfplumber)
- Requires: `pdf2image`, `pytesseract`, system Tesseract installed
- Converts each page at **200 DPI**, then OCRs with `pytesseract.image_to_data()`
- Date format: `DD-MM-YYYY` (e.g. `31-03-2025`)

**Image dimensions at 200 DPI:** 1700 × 2200px

**Column X-boundaries (pixel coords):**
```python
_CB_X_DATE_MAX        = 280   # date column
_CB_X_PARTICULARS_MAX = 880   # particulars column
_CB_X_DEPOSITS_MAX    = 1140  # deposits column
_CB_X_WITHDRAWALS_MAX = 1440  # withdrawals column
# balance: x > 1440
```

**Key insight — Canara layout:**
- Particulars span multiple lines ABOVE the date (OCR reads left-to-right, top-to-bottom across columns)
- Each transaction ends with a `Chq:` line
- `_CB_CHQ_RE = re.compile(r'^chq[;:.]?$', re.IGNORECASE)` — matches Chq variants (OCR sometimes reads `Chq;:`)

**Particulars boundary logic:**
- Particulars for transaction N = text in particulars column between:
  - START: Chq line of transaction (N-1) + 30px
  - END: own Chq line (first Chq y >= date_y)
- First transaction: starts from "Particulars" header row + 20px

**Account info extraction** (`extract_canara_account_info`):
- Runs OCR on page 1 only
- Two-column header layout: labels on left column, values on right column — OCR reads them separately
- Special patterns needed for Branch Code/Name/IFSC because they appear as label-block then value-block:
```python
"Branch Code": r'Branch Code\s*\nBranch Name\s*\nIFSC Code\s*\n\nAddress\s*\nRoad\s*\n\n(\d+)'
"Branch Name": r'Branch Code\s*\nBranch Name\s*\nIFSC Code\s*\n\nAddress\s*\nRoad\s*\n\n\d+\s*\n\n([A-Z]+)'
"IFSC Code":   r'(CNRB\w+|[A-Z]{4}0\w+)'
```

**Processing time:** ~3–5 seconds per page (OCR overhead). 5-page ePassbook = 15–25 seconds total.

---

## Key Design Decisions Made

| Decision | Reason |
|---|---|
| Streamlit (not FastAPI + React) | Uncle needs simple UI, no installs, single file to share |
| No Groq/LLM in production flow | Groq free tier hits 12k TPM limit on real PDFs; deterministic parser is faster, free, offline |
| AI Prompt mode kept as fallback | When parser fails on unknown bank layout, user can paste text into ChatGPT/Claude |
| Submit button before processing | Prevents re-processing on every Streamlit rerun (widget interaction) |
| OCR for Canara | Canara ePassbook PDFs are fully image-based (0 chars extractable by pdfplumber) |
| Router pattern | Clean separation — adding a new bank = write 2 functions + 1 elif |
| `file_bytes` read once | Streamlit reruns on every interaction; read bytes once after submit, pass everywhere |

---

## Known Issues / Limitations

- **Canara: empty particulars on some rows** — OCR sometimes can't find text between adjacent Chq boundaries (page 2 of test PDF has ~3 empty rows). This is a Tesseract limitation on dense layouts.
- **UCO: Particulars was showing amounts** — fixed by using `x1 > 360` to detect numeric columns, not decimal format (UCO amounts like `908`, `25000` have no decimal places).
- **Email regex** — UCO page 1 has TWO email addresses (customer + branch). Fixed with strict email pattern that stops at first match.
- **Groq 413 error (historical)** — was hit when trying to send whole PDF text to Groq. Removed Groq from parse flow entirely; it only exists in the AI Prompt template.
- **`col3` unused** in Mode B download section — cosmetic only, doesn't break anything.

---

## Pending / Next Steps

- [ ] SBI parser (text-based, different date format `DD MMM YYYY`, different column layout)
- [ ] HDFC parser
- [ ] `safe_float()` helper for amount summation (current `.replace(",","")` breaks on multi-comma numbers)
- [ ] Page count without double-reading PDF in Mode A (currently calls `extract_pdf_text` just for `page_count`)
- [ ] Deployment to Streamlit Community Cloud for uncle to access via URL

---

## File Structure

```
STREAMLIT_PDF_TEXT/
├── app.py              ← your main streamlit file (what we've been calling streamlit_app.py)
├── requirements.txt
├── README.md
├── .gitignore
├── .streamlit/         ← streamlit config (secrets, config.toml)
├── .agents/            ← probably cursor/copilot agents config
└── __pycache__/
```

The FastAPI + React version was built first but replaced by the Streamlit app for simplicity.

---

## Test PDFs Used

| Bank | File | Type | Pages | Transactions |
|---|---|---|---|---|
| UCO Bank | `Account_statement_20260615190904.pdf` | Text-based | 16 | 500 |
| Canara Bank | `canara_epassbook.pdf` | Image-based (scanned) | 5 | 22 |