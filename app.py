"""
Capacitance Probe Tools — Unified Streamlit App
=================================================
Single Streamlit app with sidebar navigation for all probe tools:
  1. Landing page (workflow overview)
  2. CSV → Excel converter
  3. FScan verification plot
  4. Probe overlay

Domino:  app.sh → streamlit run app.py --server.port 8888
Local:   streamlit run app.py
"""

import io
import os
import re
import traceback
from datetime import time as dt_time
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ── Page config (must be first Streamlit call) ───────────────────────────
st.set_page_config(
    page_title="Capacitance Probe Tools",
    page_icon="📊",
    layout="wide",
)

# ═════════════════════════════════════════════════════════════════════════
# Shared limit definitions
# ═════════════════════════════════════════════════════════════════════════

VENDOR_LIMITS = {
    300:  (16.00, -16.00),  374:  (12.23, -12.23),  466:  ( 9.79, -9.79),
    580:  ( 8.19,  -8.19),  720:  ( 7.15,  -7.15),  896:  ( 6.46, -6.46),
    897:  ( 6.46,  -6.46), 1118:  ( 5.99,  -5.99), 1392:  ( 5.68, -5.68),
   1729:  ( 5.47,  -5.47), 2155:  ( 5.33,  -5.33), 2689:  ( 5.24, -5.24),
   3347:  ( 5.17,  -5.17), 4158:  ( 5.12,  -5.12), 5188:  ( 5.09, -5.09),
   6447:  ( 5.07,  -5.07), 8030:  ( 5.05,  -5.05), 9995:  ( 5.04, -5.04),
}

STATE0_LIMITS_GOOD = {
    300: (11.59, -8.11),  374: ( 8.53, -5.97),  466: ( 6.23, -4.36),
    580: ( 4.63, -3.24),  720: ( 3.99, -2.79),  896: ( 3.80, -2.66),
   1118: ( 3.69, -2.58), 1392: ( 3.58, -2.51), 1729: ( 3.50, -2.45),
   2155: ( 3.45, -2.42), 2689: ( 3.42, -2.39), 3347: ( 3.37, -2.36),
   4158: ( 3.34, -2.34), 5188: ( 3.29, -2.30), 6447: ( 3.27, -2.29),
   8030: ( 3.25, -2.28), 9995: ( 3.23, -2.26),
}

STATE0_LIMITS_OKAY = {
    300: (15.07, -10.55),  374: (11.09, -7.76),  466: ( 8.32, -5.82),
    580: ( 6.96,  -4.87),  720: ( 6.08, -4.26),  896: ( 5.28, -3.70),
   1118: ( 4.97,  -3.48), 1392: ( 4.72, -3.30), 1729: ( 4.55, -3.19),
   2155: ( 4.48,  -3.14), 2689: ( 4.44, -3.11), 3347: ( 4.37, -3.06),
   4158: ( 4.34,  -3.04), 5188: ( 4.27, -2.99), 6447: ( 4.25, -2.98),
   8030: ( 4.22,  -2.95), 9995: ( 4.20, -2.94),
}

STATE3_LIMITS_GOOD = {
    300: (10.34, -7.24),  374: ( 7.45, -5.22),  466: ( 5.54, -3.88),
    580: ( 4.38, -3.07),  720: ( 4.05, -2.83),  896: ( 3.81, -2.67),
   1118: ( 3.70, -2.59), 1392: ( 3.48, -2.44), 1729: ( 3.39, -2.37),
   2155: ( 3.32, -2.32), 2689: ( 3.24, -2.27), 3347: ( 3.18, -2.23),
   4158: ( 3.13, -2.19), 5188: ( 3.09, -2.16), 6447: ( 3.03, -2.12),
   8030: ( 2.99, -2.09), 9995: ( 2.94, -2.06),
}

STATE3_LIMITS_OKAY = {
    300: (13.44, -9.41),  374: ( 9.69, -6.78),  466: ( 7.20, -5.04),
    580: ( 6.10, -4.27),  720: ( 5.62, -3.93),  896: ( 5.26, -3.68),
   1118: ( 4.85, -3.39), 1392: ( 4.60, -3.22), 1729: ( 4.41, -3.09),
   2155: ( 4.31, -3.02), 2689: ( 4.21, -2.95), 3347: ( 4.14, -2.90),
   4158: ( 4.07, -2.85), 5188: ( 4.02, -2.81), 6447: ( 3.94, -2.76),
   8030: ( 3.89, -2.72), 9995: ( 3.82, -2.67),
}

LIMIT_SETS = {
    "VendorLimits":      (VENDOR_LIMITS,      "red",     "dash"),
    "NewLimits-S0 Good": (STATE0_LIMITS_GOOD,  "orange",  "dashdot"),
    "NewLimits-S0 Okay": (STATE0_LIMITS_OKAY,  "#e65100", "dot"),
    "NewLimits-S3 Good": (STATE3_LIMITS_GOOD,  "#1b5e20", "dashdot"),
    "NewLimits-S3 Okay": (STATE3_LIMITS_OKAY,  "#6a1b9a", "dot"),
}

TRACE_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
    "#393b79", "#637939", "#8c6d31", "#843c39",
]

CSV_TO_STANDARD_FREQ = {896: 897}


# ═════════════════════════════════════════════════════════════════════════
# Shared parsers
# ═════════════════════════════════════════════════════════════════════════

def parse_fscan_bytes(file_bytes, filename):
    """Parse Excel verification file from in-memory bytes."""
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

    metadata = {"filename": filename}
    for idx in range(min(15, len(df))):
        cell = str(df.iloc[idx, 0]) if pd.notna(df.iloc[idx, 0]) else ""
        if "Batch name:" in cell:
            metadata["batch_name"] = cell.split("Batch name:")[-1].strip()
        elif "Creation Date:" in cell:
            metadata["creation_date"] = cell.split("Creation Date:")[-1].strip()
        elif "Sensor serial number:" in cell:
            metadata["sensor_serial"] = cell.split("Sensor serial number:")[-1].strip()
        elif "Status:" in cell:
            metadata["status"] = cell.split("Status:")[-1].strip()

    header_row = None
    for idx in range(len(df)):
        cell = str(df.iloc[idx, 0]) if pd.notna(df.iloc[idx, 0]) else ""
        if cell.strip() == "Date":
            header_row = idx
            break

    if header_row is None:
        raise ValueError(f"Could not find data header row in {filename}")

    df_data = pd.read_excel(xls, sheet_name=raw_sheet, header=header_row)

    freq_cols = {}
    for col in df_data.columns:
        match = re.match(r"C\((\d+)kHz\)", str(col))
        if match:
            freq_cols[int(match.group(1))] = col

    if not freq_cols:
        raise ValueError(f"No frequency columns found in {filename}")

    if "Status" in df_data.columns:
        status_vals = df_data["Status"].astype(str).str.strip()
        df_meas = df_data[status_vals.isin(["OK", "Warning"])].copy()
    else:
        df_meas = df_data[df_data["Date"].notna()].copy()

    freq_averages = {}
    for freq_khz, col in sorted(freq_cols.items()):
        values = pd.to_numeric(df_meas[col], errors="coerce").dropna()
        if len(values) > 0:
            freq_averages[freq_khz] = values.mean()

    metadata["n_measurements"] = len(df_meas)
    return metadata, freq_averages


# ═════════════════════════════════════════════════════════════════════════
# CSV-to-Excel conversion functions
# ═════════════════════════════════════════════════════════════════════════

def parse_csv_sections(file_bytes, filename):
    text = None
    for enc in ("utf-8-sig", "latin-1", "cp1252"):
        try:
            text = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"Could not decode {filename}")

    lines = [l.strip() for l in text.splitlines() if not l.strip().startswith("sep=")]
    section = None
    header_lines, event_lines = [], []
    measure_header = None
    measure_data_lines = []

    for line in lines:
        if line == "[[HEADER]]":   section = "header"; continue
        elif line == "[[EVENTS]]": section = "events"; continue
        elif line == "[[MEASURE]]": section = "measure"; continue
        if not line: continue
        if section == "header":   header_lines.append(line)
        elif section == "events": event_lines.append(line)
        elif section == "measure":
            if measure_header is None: measure_header = line
            else: measure_data_lines.append(line)

    return header_lines, event_lines, measure_header, measure_data_lines


def parse_excel_sections(file_bytes, filename):
    buf = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(buf)
    raw_sheet = None
    for name in xls.sheet_names:
        if "Log Data" in name: raw_sheet = name; break
    if raw_sheet is None: raw_sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=raw_sheet, header=None)

    section = None
    header_lines, event_lines = [], []
    measure_header = None
    measure_data_lines = []

    for idx in range(len(df)):
        cell0 = str(df.iloc[idx, 0]).strip() if pd.notna(df.iloc[idx, 0]) else ""
        if cell0 == "[[HEADER]]":   section = "header"; continue
        elif cell0 == "[[EVENTS]]": section = "events"; continue
        elif cell0 == "[[MEASURE]]": section = "measure"; continue
        if not cell0 and section != "measure": continue

        if section == "header": header_lines.append(cell0)
        elif section == "events":
            col1 = str(df.iloc[idx, 1]).strip() if df.shape[1] > 1 and pd.notna(df.iloc[idx, 1]) else ""
            event_lines.append(f"{cell0},{col1}" if col1 else cell0)
        elif section == "measure":
            if cell0 == "Date":
                cols = [str(df.iloc[idx, c]).strip() if pd.notna(df.iloc[idx, c]) else "" for c in range(df.shape[1])]
                measure_header = ",".join(cols)
            elif cell0 and cell0 != "nan":
                vals = []
                for c in range(df.shape[1]):
                    v = df.iloc[idx, c]
                    vals.append("" if pd.isna(v) else str(v))
                measure_data_lines.append(",".join(vals))

    return header_lines, event_lines, measure_header, measure_data_lines


def parse_input_sections(file_bytes, filename):
    suffix = Path(filename).suffix.lower()
    if suffix == ".csv":
        return parse_csv_sections(file_bytes, filename)
    elif suffix in (".xlsx", ".xls"):
        return parse_excel_sections(file_bytes, filename)
    else:
        raise ValueError(f"Unsupported format '{suffix}'")


def parse_record_time(val):
    parts = str(val).strip().split(":")
    if len(parts) == 3:
        return dt_time(int(parts[0]), int(parts[1]), int(parts[2]))
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
    def pad(values):
        row = list(values) + [""] * (n_cols - len(values))
        return row[:n_cols]

    rows = [pad(["[[HEADER]]"])]
    for line in header_lines: rows.append(pad([line]))
    rows.append(pad([]))
    rows.append(pad(["[[EVENTS]]"]))
    for line in event_lines: rows.append(pad(line.split(",", 1)))
    rows.append(pad([]))
    rows.append(pad(["[[MEASURE]]"]))
    rows.append(columns)

    data_rows = []
    for line in measure_data_lines:
        values = line.split(",")
        parsed = []
        for i, val in enumerate(values):
            val = val.strip()
            col_name = columns[i] if i < len(columns) else ""
            if col_name == "Date": parsed.append(val)
            elif col_name == "Record Time": parsed.append(parse_record_time(val))
            elif col_name == "Culture Time": parsed.append(None if val == "" else val)
            elif col_name == "Status": parsed.append(val)
            else:
                try: parsed.append(float(val)) if val != "" else parsed.append(None)
                except ValueError: parsed.append(val)
        data_rows.append(pad(parsed))

    rows.extend(data_rows)

    avg_row = [None] * n_cols
    for ci in range(n_cols):
        if columns[ci] in ("Date", "Record Time", "Culture Time", "Status"): continue
        nums = [dr[ci] for dr in data_rows if isinstance(dr[ci], (int, float)) and not (isinstance(dr[ci], float) and np.isnan(dr[ci]))]
        if nums: avg_row[ci] = np.mean(nums)
    rows.append(avg_row)
    rows.append([np.nan] * n_cols)
    rows.append([np.nan] * n_cols)

    return columns, rows, data_rows


def build_verification_sheet(columns, data_rows, metadata):
    freq_cols = {}
    for i, col in enumerate(columns):
        m = re.match(r"C\((\d+)kHz\)", col)
        if m:
            fk = int(m.group(1))
            freq_cols[CSV_TO_STANDARD_FREQ.get(fk, fk)] = i

    freq_avg = {}
    for fk, ci in sorted(freq_cols.items()):
        vals = [dr[ci] for dr in data_rows if isinstance(dr[ci], (int, float)) and not (isinstance(dr[ci], float) and np.isnan(dr[ci]))]
        if vals: freq_avg[fk] = np.mean(vals)

    batch = metadata.get("Batch name", "Probe")
    label = "BRX" + re.sub(r"[^0-9]", "", batch.split("_")[0]) if batch else "Probe"

    vrows = [[None]*5, [None, "Frequency", label, "Verification Limit +", "Verification Limit -"]]
    for fk in sorted(freq_avg):
        lim = VENDOR_LIMITS.get(fk, (None, None))
        vrows.append([None, fk, freq_avg[fk], lim[0], lim[1]])
    return vrows


def convert_bytes_to_excel(file_bytes, filename, custom_output_name=None):
    hl, el, mh, mdl = parse_input_sections(file_bytes, filename)
    metadata = build_metadata_dict(hl)
    columns, all_rows, data_rows = build_log_data_sheet(hl, el, mh, mdl)
    vrows = build_verification_sheet(columns, data_rows, metadata)

    cd = metadata.get("Creation Date", "")
    if cd:
        parts = cd.replace(".", " ").replace(":", " ").split()
        sheet_name = f"{parts[2]}-{parts[1]}-{parts[0]}_{parts[3]}_{parts[4]}_{parts[5]}_Log Data" if len(parts) >= 6 else Path(filename).stem
    else:
        sheet_name = Path(filename).stem

    if custom_output_name:
        output_name = custom_output_name + ("" if custom_output_name.endswith(".xlsx") else ".xlsx")
    else:
        bn = metadata.get("Batch name", "").strip()
        sn = metadata.get("Sensor serial number", "").strip()
        if bn:
            output_name = "_".join([bn] + ([f"SN{sn}"] if sn else [])) + ".xlsx"
        else:
            output_name = Path(filename).stem + "_converted.xlsx"

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(all_rows).to_excel(writer, sheet_name=sheet_name[:31], index=False, header=False)
        pd.DataFrame(vrows).to_excel(writer, sheet_name="Verification Plotted", index=False, header=False)
    buf.seek(0)

    summary = {"columns": len(columns), "data_rows": len(data_rows),
                "total_rows": len(all_rows), "verification_entries": len(vrows)-2,
                "sheet_name": sheet_name[:31]}
    return output_name, buf.getvalue(), metadata, summary


# ═════════════════════════════════════════════════════════════════════════
# Shared plot helpers
# ═════════════════════════════════════════════════════════════════════════

def add_limit_traces(fig, all_freqs, selected_limits, use_abs=False):
    for limit_name in selected_limits:
        if limit_name not in LIMIT_SETS: continue
        ld, color, dash = LIMIT_SETS[limit_name]
        lf = [f for f in all_freqs if f in ld]
        if not lf: continue
        upper = [ld[f][0] for f in lf]
        lower = [ld[f][1] for f in lf]
        if use_abs:
            fig.add_trace(go.Scatter(x=lf, y=upper, mode="lines",
                line=dict(color=color, width=2, dash=dash), name=limit_name))
        else:
            fig.add_trace(go.Scatter(x=lf, y=upper, mode="lines",
                line=dict(color=color, width=2, dash=dash), name=f"{limit_name} (upper)"))
            fig.add_trace(go.Scatter(x=lf, y=lower, mode="lines",
                line=dict(color=color, width=2, dash=dash), name=f"{limit_name} (lower)", showlegend=False))


def add_data_traces(fig, all_results, use_abs=False):
    for i, (meta, freq_avg) in enumerate(all_results):
        freqs = sorted(freq_avg.keys())
        values = [abs(freq_avg[f]) if use_abs else freq_avg[f] for f in freqs]
        color = TRACE_COLORS[i % len(TRACE_COLORS)]
        label = Path(meta.get("filename", "Unknown")).stem
        sensor = meta.get("sensor_serial", "")
        n = meta.get("n_measurements", "?")
        legend = f"{label} — S/N {sensor} (n={n})" if sensor else f"{label} (n={n})"
        hover = [f"<b>{label}</b><br>Freq: {f} kHz<br>Cap: {v:.4f} pF/cm" for f, v in zip(freqs, values)]
        fig.add_trace(go.Scatter(x=freqs, y=values, mode="lines+markers",
            line=dict(color=color, width=2.5), marker=dict(size=7),
            name=legend, hovertext=hover, hoverinfo="text"))


def format_figure(fig, all_freqs, title, use_abs=False):
    fig.update_layout(
        xaxis=dict(title="Frequency (kHz)", type="log",
                   tickvals=all_freqs, ticktext=[str(f) for f in all_freqs], tickangle=45),
        yaxis=dict(title="|Capacitance| (pF/cm)" if use_abs else "Capacitance (pF/cm)"),
        title=dict(text=title, font=dict(size=18)),
        legend=dict(x=1.02, y=1, font=dict(size=10)),
        hovermode="closest", template="plotly_white",
        margin=dict(l=60, r=300, t=60, b=80), height=650)


def build_pass_fail_table(all_results):
    rows = []
    for meta, freq_avg in all_results:
        label = Path(meta.get("filename", "Unknown")).stem
        sensor = meta.get("sensor_serial", "")
        n = meta.get("n_measurements", "?")
        for freq in sorted(freq_avg.keys()):
            val = freq_avg[freq]
            def chk(lim_dict):
                u, l = lim_dict.get(freq, (None, None))
                if u is None: return "N/A"
                return "PASS" if l <= val <= u else "FAIL"
            rows.append({"File": label, "Sensor": sensor, "Scans": n,
                         "Freq (kHz)": freq, "Avg Cap": round(val, 4),
                         "VendorLimits": chk(VENDOR_LIMITS),
                         "NewLimits-S0": chk(STATE0_LIMITS_GOOD),
                         "NewLimits-S3": chk(STATE3_LIMITS_GOOD)})
    return pd.DataFrame(rows)


# ═════════════════════════════════════════════════════════════════════════
# Sidebar navigation
# ═════════════════════════════════════════════════════════════════════════

st.sidebar.markdown(
    """
    <div style="text-align:center; padding:12px 0;">
        <span style="font-size:28px;">📊</span><br>
        <b style="font-size:16px;">Capacitance Probe Tools</b>
    </div>
    """, unsafe_allow_html=True)

page = st.sidebar.radio(
    "Navigate",
    ["🏠 Home", "1️⃣ Convert (CSV → Excel)", "2️⃣ Verification Plot", "3️⃣ Probe Overlay"],
    index=0,
)


# ═════════════════════════════════════════════════════════════════════════
# PAGE: Home / Landing
# ═════════════════════════════════════════════════════════════════════════

if page == "🏠 Home":
    st.markdown(
        """
        <div style="background:#C028B9;
             color:white; padding:24px 32px; border-radius:12px; margin-bottom:24px;">
            <h1 style="margin:0; font-size:28px; color:white !important;">Capacitance Probe Verification Tools</h1>
            <p style="margin:6px 0 0 0; opacity:0.85; font-size:14px; color:white !important;">
                End-to-end workflow for converting, verifying, and tracking
                capacitance probe frequency scan data
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("## Workflow")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            """
            <div style="background:#e3f2fd; border-radius:10px; padding:20px;
                 border-left:4px solid #1565c0; min-height:260px;">
                <h3 style="color:#1565c0; margin-top:0;">Step 1: Convert</h3>
                <p style="font-size:28px; margin:8px 0;">📁 → 📊</p>
                <p><b>CSV → Model-Ready Excel</b></p>
                <p style="font-size:13px; color:#555;">
                    Upload raw CSV or Excel log data exports from the instrument.
                    The tool parses <code>[[HEADER]]</code>, <code>[[EVENTS]]</code>,
                    and <code>[[MEASURE]]</code> sections and produces a standardised
                    Excel workbook with <b>Log Data</b> and <b>Verification Plotted</b> sheets.
                </p>
            </div>
            """, unsafe_allow_html=True)

    with col2:
        st.markdown(
            """
            <div style="background:#e8f5e9; border-radius:10px; padding:20px;
                 border-left:4px solid #2e7d32; min-height:260px;">
                <h3 style="color:#2e7d32; margin-top:0;">Step 2: Verify</h3>
                <p style="font-size:28px; margin:8px 0;">📊 → 📈</p>
                <p><b>Plot Against Limits</b></p>
                <p style="font-size:13px; color:#555;">
                    Upload the model-ready Excel file and plot the frequency scan
                    against <b>3 limit sets</b>:<br>
                    • <span style="color:red;">VendorLimits</span><br>
                    • <span style="color:orange;">NewLimits-S0</span> (Good / Okay)<br>
                    • <span style="color:#1b5e20;">NewLimits-S3</span> (Good / Okay)
                </p>
            </div>
            """, unsafe_allow_html=True)

    with col3:
        st.markdown(
            """
            <div style="background:#fff3e0; border-radius:10px; padding:20px;
                 border-left:4px solid #e65100; min-height:260px;">
                <h3 style="color:#e65100; margin-top:0;">Step 3: Overlay</h3>
                <p style="font-size:28px; margin:8px 0;">📈📈📈</p>
                <p><b>Track Over Time</b></p>
                <p style="font-size:13px; color:#555;">
                    Upload multiple Excel files for the <b>same probe</b>
                    (different test dates) and overlay all results on a single
                    interactive chart. Monitor probe drift and degradation
                    over repeated verification tests.
                </p>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Quick Start")
    st.markdown(
        """
        1. Use the **sidebar** on the left to navigate between tools
        2. Start with **Step 1** if you have raw CSV log files
        3. If you already have model-ready Excel files, skip to **Step 2** or **Step 3**
        """)

    st.markdown("### Limit Sets Reference")
    ref_data = []
    for freq in sorted(VENDOR_LIMITS.keys()):
        if freq == 897: continue  # skip duplicate
        row = {"Freq (kHz)": freq}
        row["VendorLimits ±"] = f"±{VENDOR_LIMITS[freq][0]:.2f}"
        s0 = STATE0_LIMITS_GOOD.get(freq)
        if s0: row["NewLimits-S0 Good"] = f"+{s0[0]:.2f} / {s0[1]:.2f}"
        s3 = STATE3_LIMITS_GOOD.get(freq)
        if s3: row["NewLimits-S3 Good"] = f"+{s3[0]:.2f} / {s3[1]:.2f}"
        ref_data.append(row)

    with st.expander("View all limit values"):
        st.dataframe(pd.DataFrame(ref_data), use_container_width=True, height=400)


# ═════════════════════════════════════════════════════════════════════════
# PAGE: Convert (CSV → Excel)
# ═════════════════════════════════════════════════════════════════════════

elif page == "1️⃣ Convert (CSV → Excel)":
    st.markdown(
        """
        <div style="background:linear-gradient(135deg,#1565c0,#1976d2);
             color:white; padding:16px 24px; border-radius:10px; margin-bottom:20px;">
            <h2 style="margin:0; color:white !important;">Step 1: Convert Raw Log Data → Model-Ready Excel</h2>
            <p style="margin:4px 0 0 0; opacity:0.85; font-size:13px;">
                Upload CSV or Excel log data exports from the instrument
            </p>
        </div>
        """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Upload log data files",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="convert_upload",
        help="Accepts raw CSV exports (with [[HEADER]]/[[EVENTS]]/[[MEASURE]] sections) or raw Excel log files.",
    )

    if uploaded_files:
        st.divider()

        # ── Convert all files ──
        results = []
        progress_bar = st.progress(0)

        for i, uf in enumerate(uploaded_files):
            try:
                file_bytes = uf.getvalue()       # always returns full content
                out_name, excel_bytes, meta, summary = convert_bytes_to_excel(
                    file_bytes, uf.name
                )
                results.append({
                    "status": "ok", "input": uf.name,
                    "output_name": out_name, "excel_bytes": excel_bytes,
                    "metadata": meta, "summary": summary,
                })
            except Exception as e:
                results.append({
                    "status": "error", "input": uf.name,
                    "error": str(e), "traceback": traceback.format_exc(),
                })
            progress_bar.progress((i + 1) / len(uploaded_files))

        ok_results = [r for r in results if r["status"] == "ok"]
        err_results = [r for r in results if r["status"] == "error"]

        # ── Auto-save to output folder ──
        try:
            output_dir = Path(os.environ.get("DOMINO_WORKING_DIR", ".")) / "output"
            output_dir.mkdir(parents=True, exist_ok=True)
            for r in ok_results:
                (output_dir / r["output_name"]).write_bytes(r["excel_bytes"])
        except Exception:
            output_dir = None

        # ── Status ──
        if ok_results:
            st.success(f"✅ Successfully converted {len(ok_results)} file(s)")
        if err_results:
            st.error(f"❌ Failed to convert {len(err_results)} file(s)")

        # ── Download buttons (always visible, top-level) ──
        if ok_results:
            st.subheader("Download Converted Files")
            for idx, r in enumerate(ok_results):
                st.download_button(
                    label=f"⬇ Download {r['output_name']}",
                    data=r["excel_bytes"],
                    file_name=r["output_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{idx}_{r['input']}",
                )

            if output_dir:
                st.caption(f"Files also saved to: `{output_dir}`")

        # ── Details per file ──
        for idx, r in enumerate(ok_results):
            with st.expander(f"✅  {r['input']}  →  {r['output_name']}"):
                meta_lines = [f"**{k}:** {v}" for k, v in r["metadata"].items()]
                st.markdown("  \n".join(meta_lines))

                s = r["summary"]
                st.markdown(
                    f"**Columns:** {s['columns']} · "
                    f"**Data rows:** {s['data_rows']} · "
                    f"**Sheet:** {s['sheet_name']} · "
                    f"**Verification entries:** {s['verification_entries']}"
                )

        # ── Error details ──
        for r in err_results:
            with st.expander(f"❌  {r['input']}  — ERROR", expanded=True):
                st.error(r["error"])
                if r.get("traceback"):
                    st.code(r["traceback"], language="python")

        if ok_results:
            st.info("Proceed to **Step 2** (Verification Plot) using the sidebar.")
    else:
        st.info("👆 Upload one or more raw log data files to begin.")
        with st.expander("ℹ️  Supported formats"):
            st.markdown(
                """
                **CSV** with `[[HEADER]]`, `[[EVENTS]]`, `[[MEASURE]]` sections  
                **Excel** (.xlsx/.xls) with the same layout in the first/Log Data sheet
                """)


# ═════════════════════════════════════════════════════════════════════════
# PAGE: Verification Plot
# ═════════════════════════════════════════════════════════════════════════

elif page == "2️⃣ Verification Plot":
    st.markdown(
        """
        <div style="background:linear-gradient(135deg,#2e7d32,#388e3c);
             color:white; padding:16px 24px; border-radius:10px; margin-bottom:20px;">
            <h2 style="margin:0; color:white !important;">Step 2: FScan Verification Plot</h2>
            <p style="margin:4px 0 0 0; opacity:0.85; font-size:13px;">
                Plot probe data against VendorLimits + NewLimits-S0 + NewLimits-S3
            </p>
        </div>
        """, unsafe_allow_html=True)

    selected_limits = st.multiselect(
        "Limit sets to display",
        options=list(LIMIT_SETS.keys()),
        default=["VendorLimits", "NewLimits-S0 Good", "NewLimits-S3 Good"],
        key="verify_limits",
    )
    use_abs = st.checkbox("Absolute values", key="verify_abs")

    uploaded_files = st.file_uploader(
        "Upload model-ready Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="verify_upload",
        help="Upload the Excel files produced by Step 1 (or any Excel with Log Data sheet).",
    )

    if uploaded_files:
        all_results, errors = [], []
        for uf in uploaded_files:
            try:
                meta, fa = parse_fscan_bytes(uf.read(), uf.name)
                all_results.append((meta, fa))
            except Exception as e:
                errors.append(f"{uf.name}: {e}")

        if errors:
            for e in errors: st.error(e)
        if not all_results:
            st.warning("No valid files."); st.stop()

        all_freqs = sorted(set(f for _, fa in all_results for f in fa))

        fig = go.Figure()
        add_limit_traces(fig, all_freqs, selected_limits, use_abs)
        add_data_traces(fig, all_results, use_abs)
        format_figure(fig, all_freqs, "FScan Verification" + (" — Absolute" if use_abs else ""), use_abs)

        # Pass/fail annotation
        overall_pass = True
        for _, fa in all_results:
            for f, val in fa.items():
                if f in VENDOR_LIMITS:
                    u, l = VENDOR_LIMITS[f]
                    if val > u or val < l: overall_pass = False
        fig.add_annotation(x=0.99, y=0.99, xref="paper", yref="paper",
            text=f"<b>VendorLimits: {'ALL PASS' if overall_pass else 'FAIL DETECTED'}</b>",
            showarrow=False, font=dict(size=14, color="white"),
            bgcolor="green" if overall_pass else "red", borderpad=6, opacity=0.9,
            xanchor="right", yanchor="top")

        st.plotly_chart(fig, use_container_width=True)

        # File summary
        with st.expander("File Summary"):
            for meta, _ in all_results:
                st.text(f"  {meta.get('filename','?')}  |  S/N {meta.get('sensor_serial','N/A')}  |  {meta.get('creation_date','N/A')}  |  n={meta.get('n_measurements','?')}")

        with st.expander("Data Table"):
            st.dataframe(build_pass_fail_table(all_results), use_container_width=True, height=400)

        st.info("✅ Done! For repeated tests of the same probe, proceed to **Step 3** (Overlay).")
    else:
        st.info("👆 Upload one or more model-ready Excel files (from Step 1).")


# ═════════════════════════════════════════════════════════════════════════
# PAGE: Probe Overlay
# ═════════════════════════════════════════════════════════════════════════

elif page == "3️⃣ Probe Overlay":
    st.markdown(
        """
        <div style="background:linear-gradient(135deg,#e65100,#f57c00);
             color:white; padding:16px 24px; border-radius:10px; margin-bottom:20px;">
            <h2 style="margin:0; color:white !important;">Step 3: Probe Overlay — Verification Over Time</h2>
            <p style="margin:4px 0 0 0; opacity:0.85; font-size:13px;">
                Upload multiple test files for the same probe to track drift
            </p>
        </div>
        """, unsafe_allow_html=True)

    probe_label = st.text_input("Probe label (optional)", placeholder="e.g. Probe 360",
                                 help="Used in the plot title. Auto-detected if blank.", key="overlay_label")
    selected_limits = st.multiselect(
        "Limit sets to display",
        options=list(LIMIT_SETS.keys()),
        default=["VendorLimits", "NewLimits-S0 Good", "NewLimits-S3 Good"],
        key="overlay_limits",
    )
    use_abs = st.checkbox("Absolute values", key="overlay_abs")

    uploaded_files = st.file_uploader(
        "Upload Excel files (same probe, different dates)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="overlay_upload",
        help="Upload model-ready Excel files for one probe — each from a different test date.",
    )

    if uploaded_files:
        all_results, errors = [], []
        for uf in uploaded_files:
            try:
                meta, fa = parse_fscan_bytes(uf.read(), uf.name)
                all_results.append((meta, fa))
            except Exception as e:
                errors.append(f"{uf.name}: {e}")

        if errors:
            for e in errors: st.error(e)
        if not all_results:
            st.warning("No valid files."); st.stop()

        # Auto-detect probe label
        if not probe_label:
            batches = [m.get("batch_name", "") for m, _ in all_results if m.get("batch_name")]
            if batches:
                match = re.match(r"(\d+)", batches[0])
                probe_label = f"Probe {match.group(1)}" if match else batches[0]
            else:
                probe_label = "Probe"

        with st.expander(f"File Summary — {len(all_results)} file(s)", expanded=False):
            rows = [{"File": m.get("filename","?"), "S/N": m.get("sensor_serial","N/A"),
                     "Batch": m.get("batch_name","N/A"), "Date": m.get("creation_date","N/A"),
                     "Scans": m.get("n_measurements","?")} for m, _ in all_results]
            st.dataframe(pd.DataFrame(rows), use_container_width=True)

        all_freqs = sorted(set(f for _, fa in all_results for f in fa))

        fig = go.Figure()
        add_limit_traces(fig, all_freqs, selected_limits, use_abs)

        # Data traces with date-based labels
        for i, (meta, freq_avg) in enumerate(all_results):
            freqs = sorted(freq_avg.keys())
            values = [abs(freq_avg[f]) if use_abs else freq_avg[f] for f in freqs]
            color = TRACE_COLORS[i % len(TRACE_COLORS)]

            date_str = meta.get("creation_date", "")
            label = Path(meta.get("filename", "Unknown")).stem
            if date_str:
                parts = date_str.replace(".", " ").replace(":", " ").split()
                if len(parts) >= 3:
                    months = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
                    try: label = f"{parts[0]}{months[int(parts[1])]}{parts[2][2:]}"
                    except (ValueError, IndexError): pass

            n = meta.get("n_measurements", "?")
            hover = [f"<b>{label}</b><br>Freq: {f} kHz<br>Cap: {v:.4f} pF/cm" for f, v in zip(freqs, values)]
            fig.add_trace(go.Scatter(x=freqs, y=values, mode="lines+markers",
                line=dict(color=color, width=2.5), marker=dict(size=7),
                name=f"{label} (n={n})", hovertext=hover, hoverinfo="text"))

        format_figure(fig, all_freqs,
                      f"{probe_label} — Verification Overlay" + (" (Absolute)" if use_abs else ""),
                      use_abs)

        st.plotly_chart(fig, use_container_width=True)

        with st.expander("Data Table", expanded=False):
            st.dataframe(build_pass_fail_table(all_results), use_container_width=True, height=400)
    else:
        st.info("👆 Upload two or more Excel files for the same probe to see the overlay.")
        with st.expander("ℹ️  How to use"):
            st.markdown(
                """
                1. **Convert** raw log data to Excel using Step 1 (if needed)
                2. **Upload** multiple Excel files — each from a different test date for the same probe
                3. The chart overlays all dates with selectable limit bands
                """)


# ── Footer ───────────────────────────────────────────────────────────────
st.sidebar.markdown("---")
st.sidebar.markdown(
    "<div style='font-size:11px; color:#999; text-align:center;'>"
    "Capacitance Probe Tools<br>MS&T Labs</div>",
    unsafe_allow_html=True)
