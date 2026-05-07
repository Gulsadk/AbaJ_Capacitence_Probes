# Capacitance Probe Apps

Web applications for capacitance probe data processing, deployed on Domino Data Lab.

## Apps

### 1. CSV-to-Excel Converter (`csv_to_excel/`)

Streamlit web app for converting raw CSV/Excel log data exports into model-ready Excel workbooks.

**Features:**
- Upload CSV (with `[[HEADER]]`/`[[EVENTS]]`/`[[MEASURE]]` sections) or raw Excel log files
- Auto-generates output filename from batch name and sensor serial number
- Produces standardised Excel with **Log Data** and **Verification Plotted** sheets
- Multi-file batch processing with individual download buttons

**Domino setup:**
- App script: `csv_to_excel/app.sh`
- Port: 8888 (Domino default)

### 2. FScan Verification (`fscan_verification/`)

Flask + Plotly web app for verifying probe frequency scan data against specification limits.

**Features:**
- Upload one or more Excel log files (`.xlsx` / `.xls`)
- Interactive Plotly verification plot with pass/fail overlay
- **3 limit sets** with checkboxes: Vendor (Hamilton), State 0 (Good/Okay), State 3 (Good/Okay)
- Data table with per-frequency pass/fail status against all limit sets
- Absolute value toggle

**Domino setup:**
- App script: `fscan_verification/app.sh`
- Port: 8888 (Domino default)

### 3. Probe Overlay (`overlay/`)

Streamlit web app for overlaying multiple verification test results of the same probe over time.

**Features:**
- Upload multiple Excel files for the same probe (different test dates)
- All tests overlaid on a single interactive Plotly chart
- Selectable limit bands: Vendor, State 0 (Good/Okay), State 3 (Good/Okay)
- Auto-detects probe label from file metadata
- Data table with pass/fail against all limit tiers

**Domino setup:**
- App script: `overlay/app.sh`
- Port: 8888 (Domino default)

## Workflow

1. **Convert** raw log data (CSV) to model-ready Excel → `csv_to_excel`
2. **Verify** a single test against vendor + custom limits → `fscan_verification`
3. **Overlay** multiple tests for the same probe over time → `overlay`

## Domino Deployment

Each app folder is self-contained with its own `app.sh`, `requirements.txt`, and `app.py`.

In Domino, when creating an App:
1. Point the Git repository to this repo
2. Set the **App script** to the path of the `app.sh` for the desired app:
   - `csv_to_excel/app.sh` for the data converter
   - `fscan_verification/app.sh` for the verification plot
   - `overlay/app.sh` for the probe overlay

## Local Development

```bash
# CSV-to-Excel Converter (Streamlit)
cd csv_to_excel
pip install -r requirements.txt
streamlit run app.py

# FScan Verification (Flask)
cd fscan_verification
pip install -r requirements.txt
python app.py
# Open http://localhost:8888

# Probe Overlay (Streamlit)
cd overlay
pip install -r requirements.txt
streamlit run app.py
```

## Repository Structure

```
capacitance-probe-apps/
├── README.md
├── .gitignore
├── csv_to_excel/
│   ├── app.py            # Streamlit web app
│   ├── app.sh            # Domino launcher
│   └── requirements.txt
├── fscan_verification/
│   ├── app.py            # Flask web app
│   ├── app.sh            # Domino launcher
│   └── requirements.txt
└── overlay/
    ├── app.py            # Streamlit web app
    ├── app.sh            # Domino launcher
    └── requirements.txt
```
