import io

import pdfplumber
import streamlit as st
import streamlit.components.v1 as components


AI_EXCEL_PROMPT_TEMPLATE = """You are an expert bank statement data extraction assistant.

Create a clean Excel-ready output from the PDF text below.

Requirements:
- Preserve every transaction row in the statement.
- Create two sheets: Account Details and Transactions.
- Account Details columns: Field, Value.
- Transactions columns: Date, Particulars, Withdrawal, Deposit, Balance.
- Keep amounts exactly as written unless a clear formatting cleanup is needed.
- Put withdrawals and deposits in separate columns. Leave missing values blank.
- If a transaction description wraps across lines, merge it into the same row.
- Ignore page headers, footers, disclaimers, duplicate column headings, and summary noise.
- Return the result as CSV blocks, one block per sheet, so I can paste/import it into Excel.
- If anything is uncertain, add a short Notes section after the CSV blocks.

PDF text:
{pdf_text}
"""


def extract_pdf_text(file_bytes: bytes) -> tuple[str, int]:
    page_texts = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        total_pages = len(pdf.pages)
        for page_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            page_texts.append(f"--- Page {page_idx} of {total_pages} ---\n{text.strip()}")

    return "\n\n".join(page_texts).strip(), len(page_texts)


def build_ai_excel_prompt(pdf_text: str) -> str:
    return AI_EXCEL_PROMPT_TEMPLATE.format(pdf_text=pdf_text.strip())


def render_copy_button(text: str) -> None:
    escaped_text = text.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")
    components.html(
        f"""
        <button id="copy-prompt" style="
            width: 100%;
            border: 0;
            border-radius: 8px;
            padding: 0.75rem 1rem;
            background: #0f1f35;
            color: white;
            font: 600 0.95rem system-ui, sans-serif;
            cursor: pointer;
        ">Copy prompt</button>
        <script>
        const button = document.getElementById("copy-prompt");
        button.addEventListener("click", async () => {{
            try {{
                await navigator.clipboard.writeText(`{escaped_text}`);
                button.textContent = "Copied";
                setTimeout(() => button.textContent = "Copy prompt", 1800);
            }} catch (error) {{
                button.textContent = "Select and copy manually";
                setTimeout(() => button.textContent = "Copy prompt", 2200);
            }}
        }});
        </script>
        """,
        height=56,
    )


st.set_page_config(
    page_title="PDF to AI Excel Prompt",
    page_icon="page",
    layout="centered",
)

st.title("PDF to AI Excel Prompt")
st.caption("Extract all text from a bank statement PDF and create a ChatGPT-ready Excel prompt.")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    file_bytes = uploaded_file.getvalue()

    with st.spinner("Extracting text and building prompt..."):
        try:
            extracted_text, page_count = extract_pdf_text(file_bytes)
        except Exception as exc:
            st.error(f"Could not extract text from this PDF: {exc}")
            st.stop()

    if not extracted_text:
        st.warning("No extractable text found in this PDF.")
        st.stop()

    prompt = build_ai_excel_prompt(extracted_text)

    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Pages", page_count)
    col_b.metric("Text chars", f"{len(extracted_text):,}")
    col_c.metric("Prompt chars", f"{len(prompt):,}")

    render_copy_button(prompt)

    st.download_button(
        "Download prompt as .txt",
        data=prompt,
        file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_ai_excel_prompt.txt",
        mime="text/plain",
        use_container_width=True,
    )

    st.text_area("ChatGPT-ready prompt", value=prompt, height=420)

    with st.expander("View extracted PDF text"):
        st.text_area("Extracted text", value=extracted_text, height=300)
else:
    st.info("Upload a PDF to generate the prompt.")
