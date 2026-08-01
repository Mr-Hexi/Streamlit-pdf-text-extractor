# Bank Statement → Excel

A Streamlit app that turns bank statement PDFs into Excel-ready data. It currently supports parsing UCO Bank statements and can also prepare raw PDF text for AI-assisted extraction.

## What the app does

1. Upload a bank statement PDF.
2. Select the bank format to use.
3. Choose one of two modes:
   - Parse & Download Excel: extracts account details and transactions, shows a preview, and downloads a styled Excel workbook.
   - Extract Text for AI: extracts raw PDF text and builds a prompt that can be pasted into ChatGPT, Claude, or Gemini for further CSV/Excel conversion.

## Current capabilities

- Upload PDF bank statements
- Extract account details such as statement period, account number, customer ID, branch details, and address for UCO Bank statements
- Extract transaction rows into columns for date, particulars, withdrawals, deposits, and balance
- Preview the parsed transactions in the app
- Download the result as an Excel file
- Download the extracted text or AI prompt as text files

## Notes and limitations

- The implemented parser currently targets UCO Bank statements.
- SBI and HDFC are listed in the UI as coming soon.
- Best results come from text-based PDFs. Scanned or image-based PDFs may not parse correctly.

## Requirements

Install the required Python packages:

```powershell
pip install -r requirements.txt
```

## Run locally

```powershell
streamlit run app.py
```

## Deployment

This app can be deployed on Streamlit Community Cloud or any other host that supports Streamlit. The main entry file is app.py.
