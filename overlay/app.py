"""
Probe Overlay Web App — Domino Data Lab Edition
================================================
Streamlit web application for overlaying multiple verification test
results of the same capacitance probe over time.

Upload multiple Excel files (different test dates for the same probe)
and see them overlaid on a single interactive plot with selectable
limit bands (Vendor, State 0, State 3).

Domino:
    - Runs on port 8888 via app.sh
    - Launched with: streamlit run app.py --server.port 8888

Local:
    streamlit run app.py
"""

import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ── Page config ──────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Probe Overlay — Verification Over Time",
    page_icon="📈",
    layout="wide",
)

# ═════════════════════════════════════════════════════════════════════════
# Limit definitions
# ═════════════════════════════════════════════════════════════════════════

VENDOR_LIMITS = {
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

STATE0_LIMITS_GOOD = {
      300: ( 11.59,   -8.11),
      374: (  8.53,   -5.97),
      466: (  6.23,   -4.36),
      580: (  4.63,   -3.24),
      720: (  3.99,   -2.79),
      896: (  3.80,   -2.66),
     1118: (  3.69,   -2.58),
     1392: (  3.58,   -2.51),
     1729: (  3.50,   -2.45),
     2155: (  3.45,   -2.42),
     2689: (  3.42,   -2.39),
     3347: (  3.37,   -2.36),
     4158: (  3.34,   -2.34),
     5188: (  3.29,   -2.30),
     6447: (  3.27,   -2.29),
     8030: (  3.25,   -2.28),
     9995: (  3.23,   -2.26),
}

STATE0_LIMITS_OKAY = {
      300: ( 15.07,  -10.55),
      374: ( 11.09,   -7.76),
      466: (  8.32,   -5.82),
      580: (  6.96,   -4.87),
      720: (  6.08,   -4.26),
      896: (  5.28,   -3.70),
     1118: (  4.97,   -3.48),
     1392: (  4.72,   -3.30),
     1729: (  4.55,   -3.19),
     2155: (  4.48,   -3.14),
     2689: (  4.44,   -3.11),
     3347: (  4.37,   -3.06),
     4158: (  4.34,   -3.04),
     5188: (  4.27,   -2.99),
     6447: (  4.25,   -2.98),
     8030: (  4.22,   -2.95),
     9995: (  4.20,   -2.94),
}

STATE3_LIMITS_GOOD = {
      300: ( 10.34,   -7.24),
      374: (  7.45,   -5.22),
      466: (  5.54,   -3.88),
      580: (  4.38,   -3.07),
      720: (  4.05,   -2.83),
      896: (  3.81,   -2.67),
     1118: (  3.70,   -2.59),
     1392: (  3.48,   -2.44),
     1729: (  3.39,   -2.37),
     2155: (  3.32,   -2.32),
     2689: (  3.24,   -2.27),
     3347: (  3.18,   -2.23),
     4158: (  3.13,   -2.19),
     5188: (  3.09,   -2.16),
     6447: (  3.03,   -2.12),
     8030: (  2.99,   -2.09),
     9995: (  2.94,   -2.06),
}

STATE3_LIMITS_OKAY = {
      300: ( 13.44,   -9.41),
      374: (  9.69,   -6.78),
      466: (  7.20,   -5.04),
      580: (  6.10,   -4.27),
      720: (  5.62,   -3.93),
      896: (  5.26,   -3.68),
     1118: (  4.85,   -3.39),
     1392: (  4.60,   -3.22),
     1729: (  4.41,   -3.09),
     2155: (  4.31,   -3.02),
     2689: (  4.21,   -2.95),
     3347: (  4.14,   -2.90),
     4158: (  4.07,   -2.85),
     5188: (  4.02,   -2.81),
     6447: (  3.94,   -2.76),
     8030: (  3.89,   -2.72),
     9995: (  3.82,   -2.67),
}

LIMIT_SETS = {
    "Vendor (Hamilton)":  (VENDOR_LIMITS,       "red",       "dash"),
    "State 0 — Good":    (STATE0_LIMITS_GOOD,  "green",     "dashdot"),
    "State 0 — Okay":    (STATE0_LIMITS_OKAY,  "orange",    "dot"),
    "State 3 — Good":    (STATE3_LIMITS_GOOD,  "#1b5e20",   "dashdot"),
    "State 3 — Okay":    (STATE3_LIMITS_OKAY,  "#e65100",   "dot"),
}

TRACE_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
    "#393b79", "#637939", "#8c6d31", "#843c39",
]


# ── Parser ───────────────────────────────────────────────────────────────

def parse_fscan_bytes(file_bytes, filename):
    """Parse from in-memory bytes (uploaded file)."""
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
    </style>
    <div class="main-header">
        <h1>Probe Overlay — Verification Over Time</h1>
        <p>Upload multiple test files for the same probe to overlay results</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Sidebar controls ──
st.sidebar.header("Settings")

probe_label = st.sidebar.text_input(
    "Probe label (optional)",
    placeholder="e.g. Probe 360",
    help="Used in the plot title. If blank, auto-detected from file metadata.",
)

selected_limits = st.sidebar.multiselect(
    "Limit sets to display",
    options=list(LIMIT_SETS.keys()),
    default=["Vendor (Hamilton)", "State 0 — Good", "State 3 — Good"],
)

use_abs = st.sidebar.checkbox("Absolute values", value=False)

# ── File upload ──
uploaded_files = st.file_uploader(
    "Upload Excel verification files (same probe, different dates)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    help="Upload the model-ready Excel files for one probe. Each file "
         "represents a different test date.",
)

if uploaded_files:
    st.divider()

    # ── Parse all files ──
    all_results = []
    errors = []

    for uf in uploaded_files:
        try:
            file_bytes = uf.read()
            meta, freq_avg = parse_fscan_bytes(file_bytes, uf.name)
            all_results.append((meta, freq_avg))
        except Exception as e:
            errors.append(f"{uf.name}: {e}")

    if errors:
        for err in errors:
            st.error(err)

    if not all_results:
        st.warning("No valid files to plot.")
        st.stop()

    # ── Auto-detect probe label ──
    if not probe_label:
        batch_names = [m.get("batch_name", "") for m, _ in all_results if m.get("batch_name")]
        if batch_names:
            # Extract probe number from batch name like "360State0_02Mar"
            match = re.match(r"(\d+)", batch_names[0])
            probe_label = f"Probe {match.group(1)}" if match else batch_names[0]
        else:
            probe_label = "Probe"

    # ── File summary ──
    with st.expander(f"File Summary — {len(all_results)} file(s)", expanded=False):
        summary_rows = []
        for meta, freq_avg in all_results:
            summary_rows.append({
                "File": meta.get("filename", "?"),
                "Sensor S/N": meta.get("sensor_serial", "N/A"),
                "Batch": meta.get("batch_name", "N/A"),
                "Date": meta.get("creation_date", "N/A"),
                "Scans": meta.get("n_measurements", "?"),
                "Frequencies": len(freq_avg),
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

    # ── Build overlay figure ──
    fig = go.Figure()

    # Collect all frequencies
    all_freqs = set()
    for _, fa in all_results:
        all_freqs.update(fa.keys())
    all_freqs = sorted(all_freqs)

    # Draw limit bands
    for limit_name in selected_limits:
        if limit_name not in LIMIT_SETS:
            continue
        limits_dict, color, dash = LIMIT_SETS[limit_name]
        limit_freqs = [f for f in all_freqs if f in limits_dict]
        if not limit_freqs:
            continue

        upper = [limits_dict[f][0] for f in limit_freqs]
        lower = [limits_dict[f][1] for f in limit_freqs]

        if use_abs:
            fig.add_trace(go.Scatter(
                x=limit_freqs, y=upper, mode="lines",
                line=dict(color=color, width=2, dash=dash),
                name=limit_name,
            ))
        else:
            fig.add_trace(go.Scatter(
                x=limit_freqs, y=upper, mode="lines",
                line=dict(color=color, width=2, dash=dash),
                name=f"{limit_name} (upper)",
            ))
            fig.add_trace(go.Scatter(
                x=limit_freqs, y=lower, mode="lines",
                line=dict(color=color, width=2, dash=dash),
                name=f"{limit_name} (lower)",
                showlegend=False,
            ))

    # Draw data traces — one per file
    for i, (meta, freq_avg) in enumerate(all_results):
        freqs = sorted(freq_avg.keys())
        values = [abs(freq_avg[f]) if use_abs else freq_avg[f] for f in freqs]
        color = TRACE_COLORS[i % len(TRACE_COLORS)]

        # Build legend label from date / filename
        date_str = meta.get("creation_date", "")
        label = Path(meta.get("filename", "Unknown")).stem
        if date_str:
            # Shorten "02.03.2026 12:47:00" → "02Mar26"
            parts = date_str.replace(".", " ").replace(":", " ").split()
            if len(parts) >= 3:
                months = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                try:
                    m_idx = int(parts[1])
                    label = f"{parts[0]}{months[m_idx]}{parts[2][2:]}"
                except (ValueError, IndexError):
                    pass

        n = meta.get("n_measurements", "?")
        legend = f"{label} (n={n})"

        hover = [f"<b>{label}</b><br>Freq: {f} kHz<br>Cap: {v:.4f} pF/cm"
                 for f, v in zip(freqs, values)]

        fig.add_trace(go.Scatter(
            x=freqs, y=values, mode="lines+markers",
            line=dict(color=color, width=2.5),
            marker=dict(size=7),
            name=legend,
            hovertext=hover, hoverinfo="text",
        ))

    fig.update_layout(
        xaxis=dict(
            title="Frequency (kHz)", type="log",
            tickvals=all_freqs,
            ticktext=[str(f) for f in all_freqs],
            tickangle=45,
        ),
        yaxis=dict(
            title="|Capacitance| (pF/cm)" if use_abs else "Capacitance (pF/cm)",
        ),
        title=dict(
            text=f"{probe_label} — Verification Overlay"
                 + (" (Absolute)" if use_abs else ""),
            font=dict(size=18),
        ),
        legend=dict(x=1.02, y=1, font=dict(size=10)),
        hovermode="closest",
        template="plotly_white",
        margin=dict(l=60, r=300, t=60, b=80),
        height=650,
    )

    st.plotly_chart(fig, use_container_width=True)

    # ── Data table ──
    with st.expander("Data Table", expanded=False):
        table_rows = []
        for meta, freq_avg in all_results:
            label = Path(meta.get("filename", "Unknown")).stem
            sensor = meta.get("sensor_serial", "")
            for freq in sorted(freq_avg.keys()):
                val = freq_avg[freq]
                v_u, v_l = VENDOR_LIMITS.get(freq, (None, None))
                vendor_ok = (v_l <= val <= v_u) if v_u is not None else None
                s0_u, s0_l = STATE0_LIMITS_GOOD.get(freq, (None, None))
                s0_ok = (s0_l <= val <= s0_u) if s0_u is not None else None
                s3_u, s3_l = STATE3_LIMITS_GOOD.get(freq, (None, None))
                s3_ok = (s3_l <= val <= s3_u) if s3_u is not None else None

                def s(v):
                    return "PASS" if v else ("FAIL" if v is False else "N/A")

                table_rows.append({
                    "File": label,
                    "Sensor": sensor,
                    "Freq (kHz)": freq,
                    "Avg Cap (pF/cm)": round(val, 4),
                    "Vendor": s(vendor_ok),
                    "State 0": s(s0_ok),
                    "State 3": s(s3_ok),
                })

        st.dataframe(pd.DataFrame(table_rows), use_container_width=True, height=400)

else:
    st.info("👆 Upload two or more Excel files for the same probe to see the overlay plot.")

    with st.expander("ℹ️  How to use"):
        st.markdown(
            """
            1. **Convert raw log data** to model-ready Excel using the
               CSV-to-Excel converter app (if needed).
            2. **Upload multiple Excel files** — each from a different
               verification test date for the same probe.
            3. The plot overlays all test dates on a single chart with
               selectable limit bands.
            4. Use the sidebar to toggle limit sets and absolute values.
            """
        )

st.markdown(
    "<div style='text-align:center; color:#999; font-size:11px; "
    "margin-top:40px; border-top:1px solid #eee; padding-top:12px;'>"
    "Probe Overlay — Capacitance Probe Verification Over Time"
    "</div>",
    unsafe_allow_html=True,
)
