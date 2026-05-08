"""
CSV / Excel Log Data → Model-Ready Excel Converter
===================================================
Streamlit web application for converting capacitance probe raw log data
exports (CSV or Excel) into the standardised Excel format expected by
the analysis and modelling scripts.

Domino:
    - Runs on port 8888 via app.sh
    - Launched with: streamlit run app.py --server.port 8888

Local:
    streamlit run app.py
"""

import io
import os
import re
from datetime import time as dt_time
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# ── Page config ──────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Log Data → Excel Converter",
    page_icon="📊",
    layout="wide",
)

# ── Verification limits ─────────────────────────────────────────────────
VERIFICATION_LIMITS = {
    300:  (16.00, -16.00),
    374:  (12.23, -12.23),
    466:  ( 9.79,  -9.79),
    580:  ( 8.19,  -8.19),
    720:  ( 7.15,  -7.15),
    896:  ( 6.46,  -6.46),
    897:  ( 6.46,  -6.46),
    1118: ( 5.99,  -5.99),
    1392: ( 5.68,  -5.68),
    1729: ( 5.47,  -5.47),
    2155: ( 5.33,  -5.33),
    2689: ( 5.24,  -5.24),
    3347: ( 5.17,  -5.17),
    4158: ( 5.12,  -5.12),
    5188: ( 5.09,  -5.09),
    6447: ( 5.07,  -5.07),
    8030: ( 5.05,  -5.05),
    9995: ( 5.04,  -5.04),
}

CSV_TO_STANDARD_FREQ = {896: 897}


# ═════════════════════════════════════════════════════════════════════════
# Core conversion functions
# ═════════════════════════════════════════════════════════════════════════

def parse_csv_sections(file_bytes, filename):
    """Parse raw CSV bytes into header / events / measure sections."""
    text = None
    for enc in ("utf-8-sig", "latin-1", "cp1252"):
        try:
            text = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"Could not decode {filename} with any supported encoding")

    lines = [l.strip() for l in text.splitlines() if not l.strip().startswith("sep=")]

    section = None
    header_lines, event_lines = [], []
    measure_header = None
    measure_data_lines = []

    for line in lines:
        if line == "[[HEADER]]":
            section = "header"; continue
        elif line == "[[EVENTS]]":
            section = "events"; continue
        elif line == "[[MEASURE]]":
            section = "measure"; continue
        if not line:
            continue
        if section == "header":
            header_lines.append(line)
        elif section == "events":
            event_lines.append(line)
        elif section == "measure":
            if measure_header is None:
                measure_header = line
            else:
                measure_data_lines.append(line)

    return header_lines, event_lines, measure_header, measure_data_lines


def parse_excel_sections(file_bytes, filename):
    """Parse Excel log file bytes into header / events / measure sections."""
    buf = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(buf)

    raw_sheet = None
    for name in xls.sheet_names:
        if "Log Data" in name:
            raw_sheet = name
            break
    if raw_sheet is None:
        raw_sheet = xls.sheet_names[0]

    df = pd.read_excel(xls, sheet_name=raw_sheet, header=None)

    section = None
    header_lines, event_lines = [], []
    measure_header = None
    measure_data_lines = []

    for idx in range(len(df)):
        cell0 = str(df.iloc[idx, 0]).strip() if pd.notna(df.iloc[idx, 0]) else ""

        if cell0 == "[[HEADER]]":
            section = "header"; continue
        elif cell0 == "[[EVENTS]]":
            section = "events"; continue
        elif cell0 == "[[MEASURE]]":
            section = "measure"; continue

        if not cell0 and section != "measure":
            continue

        if section == "header":
            header_lines.append(cell0)
        elif section == "events":
            col1 = str(df.iloc[idx, 1]).strip() if df.shape[1] > 1 and pd.notna(df.iloc[idx, 1]) else ""
            event_lines.append(f"{cell0},{col1}" if col1 else cell0)
        elif section == "measure":
            if cell0 == "Date":
                cols = []
                for c in range(df.shape[1]):
                    v = df.iloc[idx, c]
                    cols.append(str(v).strip() if pd.notna(v) else "")
                measure_header = ",".join(cols)
            elif cell0 and cell0 != "nan":
                vals = []
                for c in range(df.shape[1]):
                    v = df.iloc[idx, c]
                    if pd.isna(v):
                        vals.append("")
                    elif hasattr(v, "strftime"):
                        vals.append(str(v))
                    else:
                        vals.append(str(v))
                measure_data_lines.append(",".join(vals))

    return header_lines, event_lines, measure_header, measure_data_lines


def parse_input_sections(file_bytes, filename):
    """Dispatch to CSV or Excel parser based on file extension."""
    suffix = Path(filename).suffix.lower()
    if suffix == ".csv":
        return parse_csv_sections(file_bytes, filename)
    elif suffix in (".xlsx", ".xls"):
        return parse_excel_sections(file_bytes, filename)
    else:
        raise ValueError(f"Unsupported file format '{suffix}'. Use .csv, .xlsx, or .xls")


def parse_record_time(val):
    parts = str(val).strip().split(":")
    if len(parts) == 3:
        h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
        return dt_time(h, m, s)
    return val


def build_metadata_dict(header_lines):
    metadata = {}
    for line in header_lines:
        if ":" in line:
            key, _, value = line.partition(":")
            metadata[key.strip()] = value.strip()
    return metadata


def build_log_data_sheet(header_lines, event_lines, measure_header, measure_data_lines):
    columns = [c.strip() for c in measure_header.split(",")]
    n_cols = len(columns)

    def pad_row(values):
        row = list(values) + [""] * (n_cols - len(values))
        return row[:n_cols]

    rows = []
    rows.append(pad_row(["[[HEADER]]"]))
    for line in header_lines:
        rows.append(pad_row([line]))
    rows.append(pad_row([]))
    rows.append(pad_row(["[[EVENTS]]"]))
    for line in event_lines:
        parts = line.split(",", 1)
        rows.append(pad_row(parts))
    rows.append(pad_row([]))
    rows.append(pad_row(["[[MEASURE]]"]))
    rows.append(columns)

    data_rows = []
    for line in measure_data_lines:
        values = line.split(",")
        parsed = []
        for i, val in enumerate(values):
            val = val.strip()
            col_name = columns[i] if i < len(columns) else ""
            if col_name == "Date":
                parsed.append(val)
            elif col_name == "Record Time":
                parsed.append(parse_record_time(val))
            elif col_name == "Culture Time":
                parsed.append(None if val == "" else val)
            elif col_name == "Status":
                parsed.append(val)
            else:
                try:
                    parsed.append(float(val)) if val != "" else parsed.append(None)
                except ValueError:
                    parsed.append(val)
        data_rows.append(pad_row(parsed))

    rows.extend(data_rows)

    # Averages row
    avg_row = [None] * n_cols
    for col_idx in range(n_cols):
        col_name = columns[col_idx]
        if col_name in ("Date", "Record Time", "Culture Time", "Status"):
            continue
        numeric_vals = []
        for dr in data_rows:
            val = dr[col_idx]
            if isinstance(val, (int, float)) and not (isinstance(val, float) and np.isnan(val)):
                numeric_vals.append(val)
        if numeric_vals:
            avg_row[col_idx] = np.mean(numeric_vals)
    rows.append(avg_row)

    rows.append([np.nan] * n_cols)
    rows.append([np.nan] * n_cols)

    return columns, rows, data_rows


def build_verification_sheet(columns, data_rows, metadata):
    freq_cols = {}
    for i, col in enumerate(columns):
        match = re.match(r"C\((\d+)kHz\)", col)
        if match:
            freq_khz = int(match.group(1))
            freq_khz = CSV_TO_STANDARD_FREQ.get(freq_khz, freq_khz)
            freq_cols[freq_khz] = i

    freq_averages = {}
    for freq_khz, col_idx in sorted(freq_cols.items()):
        vals = []
        for dr in data_rows:
            v = dr[col_idx]
            if isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v)):
                vals.append(v)
        if vals:
            freq_averages[freq_khz] = np.mean(vals)

    batch = metadata.get("Batch name", "Probe")
    probe_label = "BRX" + re.sub(r"[^0-9]", "", batch.split("_")[0]) if batch else "Probe"

    verification_rows = []
    verification_rows.append([None, None, None, None, None])
    verification_rows.append([None, "Frequency", probe_label, "Verification Limit +", "Verification Limit -"])
    for freq_khz in sorted(freq_averages.keys()):
        avg = freq_averages[freq_khz]
        lim = VERIFICATION_LIMITS.get(freq_khz, (None, None))
        verification_rows.append([None, freq_khz, avg, lim[0], lim[1]])

    return verification_rows


def convert_bytes_to_excel(file_bytes, filename, custom_output_name=None):
    """
    Convert uploaded file bytes to model-ready Excel workbook bytes.

    Returns (output_filename, excel_bytes, metadata_dict, summary_dict).
    """
    header_lines, event_lines, measure_header, measure_data_lines = parse_input_sections(
        file_bytes, filename
    )

    metadata = build_metadata_dict(header_lines)

    columns, all_rows, data_rows = build_log_data_sheet(
        header_lines, event_lines, measure_header, measure_data_lines
    )

    verification_rows = build_verification_sheet(columns, data_rows, metadata)

    # Sheet name from creation date
    creation_date = metadata.get("Creation Date", "")
    if creation_date:
        parts = creation_date.replace(".", " ").replace(":", " ").split()
        if len(parts) >= 6:
            sheet_name = f"{parts[2]}-{parts[1]}-{parts[0]}_{parts[3]}_{parts[4]}_{parts[5]}_Log Data"
        else:
            sheet_name = Path(filename).stem
    else:
        sheet_name = Path(filename).stem

    # Output filename
    if custom_output_name:
        output_name = custom_output_name
        if not output_name.endswith(".xlsx"):
            output_name += ".xlsx"
    else:
        batch_name = metadata.get("Batch name", "").strip()
        sensor_sn = metadata.get("Sensor serial number", "").strip()
        if batch_name:
            name_parts = [batch_name]
            if sensor_sn:
                name_parts.append(f"SN{sensor_sn}")
            output_name = "_".join(name_parts) + ".xlsx"
        else:
            output_name = Path(filename).stem + "_converted.xlsx"

    # Write to bytes buffer
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_log = pd.DataFrame(all_rows)
        df_log.to_excel(writer, sheet_name=sheet_name[:31], index=False, header=False)

        df_verify = pd.DataFrame(verification_rows)
        df_verify.to_excel(writer, sheet_name="Verification Plotted", index=False, header=False)

    buf.seek(0)

    summary = {
        "columns": len(columns),
        "data_rows": len(data_rows),
        "total_rows": len(all_rows),
        "verification_entries": len(verification_rows) - 2,
        "sheet_name": sheet_name[:31],
    }

    return output_name, buf.getvalue(), metadata, summary


# ═════════════════════════════════════════════════════════════════════════
# Streamlit UI
# ═════════════════════════════════════════════════════════════════════════

st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(135deg, #1a237e, #283593);
        color: white; padding: 16px 24px; border-radius: 10px;
        margin-bottom: 24px;
    }
    .main-header h1 { font-size: 22px; margin: 0; }
    .main-header p { font-size: 13px; opacity: 0.8; margin: 4px 0 0 0; }
    .success-box {
        background: #e8f5e9; border: 1px solid #a5d6a7;
        border-radius: 8px; padding: 12px 16px; margin: 8px 0;
    }
    .meta-box {
        background: #f5f5f5; border: 1px solid #e0e0e0;
        border-radius: 8px; padding: 12px 16px; margin: 8px 0;
    }
    </style>
    <div class="main-header">
        <h1>Log Data → Model-Ready Excel Converter</h1>
        <p>Capacitance probe raw data conversion tool</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    "Upload CSV or Excel log data files exported from the capacitance probe "
    "instrument. The tool converts them into the standardised Excel format "
    "with **Log Data** and **Verification Plotted** sheets."
)

uploaded_files = st.file_uploader(
    "Upload log data files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
    help="Accepts raw CSV exports (with [[HEADER]]/[[EVENTS]]/[[MEASURE]] sections) "
         "or raw Excel log files.",
)

if uploaded_files:
    st.divider()
    st.subheader(f"Processing {len(uploaded_files)} file(s)")

    results = []
    progress_bar = st.progress(0)

    for i, uploaded_file in enumerate(uploaded_files):
        file_bytes = uploaded_file.read()
        filename = uploaded_file.name

        try:
            output_name, excel_bytes, metadata, summary = convert_bytes_to_excel(
                file_bytes, filename
            )
            results.append({
                "status": "success",
                "input": filename,
                "output_name": output_name,
                "excel_bytes": excel_bytes,
                "metadata": metadata,
                "summary": summary,
            })
        except Exception as e:
            results.append({
                "status": "error",
                "input": filename,
                "error": str(e),
            })

        progress_bar.progress((i + 1) / len(uploaded_files))

    # ── Display results ──
    success_count = sum(1 for r in results if r["status"] == "success")
    error_count = sum(1 for r in results if r["status"] == "error")

    # ── Auto-save to Domino output folder ──
    output_dir = Path(os.environ.get("DOMINO_WORKING_DIR", ".")) / "output"
    if not output_dir.exists():
        output_dir.mkdir(parents=True, exist_ok=True)

    saved_files = []
    for r in [r for r in results if r["status"] == "success"]:
        out_path = output_dir / r["output_name"]
        out_path.write_bytes(r["excel_bytes"])
        saved_files.append(str(out_path))

    if success_count > 0:
        st.success(f"Successfully converted {success_count} file(s)")
        st.info(f"Files saved to: `{output_dir}/`")
        for sf in saved_files:
            st.text(f"  📄 {sf}")
    if error_count > 0:
        st.error(f"Failed to convert {error_count} file(s)")

    for r in results:
        if r["status"] == "success":
            with st.expander(f"✅  {r['input']}  →  {r['output_name']}", expanded=True):
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Metadata**")
                    for key, val in r["metadata"].items():
                        st.text(f"  {key}: {val}")

                with col2:
                    st.markdown("**Conversion Summary**")
                    s = r["summary"]
                    st.text(f"  Columns: {s['columns']}")
                    st.text(f"  Data rows: {s['data_rows']}")
                    st.text(f"  Sheet name: {s['sheet_name']}")
                    st.text(f"  Verification entries: {s['verification_entries']}")

                st.download_button(
                    label=f"⬇ Download {r['output_name']}",
                    data=r["excel_bytes"],
                    file_name=r["output_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{r['input']}",
                )
        else:
            with st.expander(f"❌  {r['input']}  — Error", expanded=True):
                st.error(r["error"])

    # ── Batch download (if multiple successes) ──
    successful = [r for r in results if r["status"] == "success"]
    if len(successful) > 1:
        st.divider()
        st.subheader("Batch Download")
        st.info("Download individual files above, or use the buttons below.")
        for r in successful:
            st.download_button(
                label=f"⬇ {r['output_name']}",
                data=r["excel_bytes"],
                file_name=r["output_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"batch_{r['input']}",
            )
else:
    st.info("👆 Upload one or more log data files to begin conversion.")

    with st.expander("ℹ️  Supported input formats"):
        st.markdown(
            """
            **CSV files** with the standard section markers:
            ```
            [[HEADER]]
            Batch name: 360State0_02Mar
            Sensor serial number: 5511
            Creation Date: 02.03.2026 12:47:00
            ...
            [[EVENTS]]
            ...
            [[MEASURE]]
            Date,Record Time,Status,C(300kHz),C(374kHz),...
            ```

            **Excel files** (.xlsx / .xls) with the same layout in the
            first sheet (or a sheet named "Log Data").
            """
        )

    with st.expander("ℹ️  Output format"):
        st.markdown(
            """
            The output Excel workbook contains two sheets:

            1. **Log Data** — Full metadata, events, measurements, and
               computed averages row
            2. **Verification Plotted** — Average capacitance per frequency
               with verification limits (±)
            """
        )

st.markdown(
    "<div style='text-align:center; color:#999; font-size:11px; "
    "margin-top:40px; border-top:1px solid #eee; padding-top:12px;'>"
    "Log Data → Excel Converter — Capacitance Probe Data Preparation"
    "</div>",
    unsafe_allow_html=True,
)
