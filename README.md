# FAE Invoice Processor

A Streamlit app that converts **Small Compressor Sales & Rentals / FAE II** invoice PDFs into formatted Excel files.

## Deploy to Streamlit Cloud (GitHub)

1. Fork or push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) → New app
3. Select your repo, set **Main file path** to `app.py`
4. Click Deploy — done!

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What it does

- Parses invoice metadata (number, date, month)
- Extracts all line items: Name, Vendor Cross Reference, Unit Price, Delivery Fee, Add'l Charges, Line Total
- Handles negative/zero additional charges (e.g. "Unable to Sell Gas" credits)
- Outputs a formatted `.xlsx` matching the invoice layout with alternating row colors and totals section
