# PDF to AI Excel Prompt

A small Streamlit app that extracts text from a PDF and builds a copy-ready prompt for ChatGPT or another AI to create Excel-friendly CSV output.

## Features

- Upload a PDF bank statement.
- Extract text from every page with page markers.
- Build a structured Excel-generation prompt.
- Copy the prompt in the browser.
- Download the prompt as a `.txt` file.
- View the raw extracted PDF text.

## Run Locally

```powershell
cd streamlit-ai-prompt
pip install -r requirements.txt
streamlit run app.py
```

If you use conda base:

```powershell
cd streamlit-ai-prompt
conda run -n base pip install -r requirements.txt
conda run -n base streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this `streamlit-ai-prompt` folder to a GitHub repository.
2. Go to Streamlit Community Cloud.
3. Create a new app from the repo.
4. Set the main file path to:

```text
streamlit-ai-prompt/app.py
```

5. Deploy.

No API keys are required because this app only prepares the prompt. The user manually pastes the prompt into ChatGPT or another AI.
