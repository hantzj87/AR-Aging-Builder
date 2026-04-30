# AR Aging Report Generator

A Streamlit app that converts a raw invoices export into a formatted **AR Aging Summary** Excel workbook — matching the Y&S Group reporting layout.

## Features

- Upload any invoices `.xlsx` export
- Select an **As of Date** — aging buckets recalculate automatically
- Generates a two-tab Excel workbook:
  - **AR Aging Summary** — pivot by network × aging bucket, whole-dollar currency
  - **Invoice Details** — full filtered invoice list with frozen header row
- Only networks with outstanding balances appear on the summary tab
- Company name normalization (YSA 2/3 → YSA, YS Tickets Spec → YS Tickets, etc.)

## Aging Buckets

| Bucket | Days Outstanding |
|---|---|
| Current | 0 or fewer |
| 1 to 30 | 1–30 days |
| 31 to 60 | 31–60 days |
| 61 to 90 | 61–90 days |
| 91 and Over | 91+ days |

## Local Setup

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_ORG/ar-aging-report.git
cd ar-aging-report

# 2. Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

The app will open at `http://localhost:8501`.

## Deploying to Streamlit Cloud

1. Push this repo to GitHub (public or private).
2. Go to [share.streamlit.io](https://share.streamlit.io) and click **New app**.
3. Select your repo, branch (`main`), and set **Main file path** to `app.py`.
4. Click **Deploy** — Streamlit Cloud installs `requirements.txt` automatically.

## Input File Requirements

The uploaded `.xlsx` must contain these columns:

| Column | Description |
|---|---|
| `Paid` | `"Yes"` / `"No"` |
| `IsCancelled` | `"Yes"` / `"No"` |
| `Bal.` | Outstanding balance (numeric) |
| `Client` | Network / marketplace name |
| `Company` | Broker entity name |
| `Inv#` | Invoice number |
| `Ext Order #` | External order reference |
| `Status` | Invoice status |
| `Created` | Invoice creation datetime |

## Project Structure

```
ar-aging-report/
├── app.py              # Streamlit UI
├── report_builder.py   # Excel generation logic
├── requirements.txt    # Python dependencies
├── .gitignore
└── README.md
```
