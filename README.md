# Ampere-Turn Engineering Calculator

This Streamlit dashboard uses the included `Amp_turn_input_sheet.xlsx` workbook as its calculation engine.

## Included behavior
- Editable fields map to the actual pink cells on `Calculations`: C8:C11, C14:C18, C21:C25, and C28.
- Existing workbook formulas are preserved and recalculated with LibreOffice headless.
- Main results come directly from `Calculations!C115` and `Calculations!C92`.
- The graph reads actual ampere-turn results from Table 03: temperatures in C140:O140, 9 V in C141:O141, and 16 V in C146:O146.
- The displayed graph is limited to -40°C through 150°C, as requested.

## Run locally
1. Install Python 3.11+ and LibreOffice.
2. Run:

```bash
pip install -r requirements.txt
streamlit run app.py
```

The local URL is normally `http://localhost:8501`.

## Deploy to a URL using Streamlit Community Cloud
1. Upload all included files to a private GitHub repository.
2. Sign in at `https://share.streamlit.io` using GitHub.
3. Select **Create app**, choose the repository, branch, and `app.py`.
4. Deploy. Streamlit will issue an HTTPS address such as `https://your-app-name.streamlit.app`.

`packages.txt` instructs Streamlit Community Cloud to install LibreOffice.
