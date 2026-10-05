# Tata 1mg Precision Validation Streamlit App

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `tata1mg_precision_app.py` and `requirements.txt`.
3. Open Streamlit Community Cloud.
4. Select the GitHub repository and `tata1mg_precision_app.py`.
5. Deploy.

## Application workflow

- Upload the precision Excel template.
- Upload one or more INTRA PDFs.
- Upload one or more INTER PDFs.
- Select Auto Detect / 2 Levels / 3 Levels.
- Analyze PDFs.
- Review analyte/level counts.
- Generate one Excel workbook per analyte.
- Download all workbooks as one ZIP.

## Precision mapping

- INTRA: up to 20 results per level.
- INTER Run 1: first 5 INTRA results.
- INTER Run 2 onward: groups of 5 results from INTER PDF data.
- Result values are used, not RLU/MAU.

Important: The app keeps an `AUTO_DATA` sheet in each generated workbook containing the extracted values. The original template's exact cell mapping may need to be configured to match your final Tata 1mg template before production use.
