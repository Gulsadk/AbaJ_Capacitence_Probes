"""
FScan Verification Web App — Domino Data Lab Edition
=====================================================
Flask + Plotly web application for capacitance probe fscan verification.

Shows uploaded Excel probe data against three limit sets:
  - Vendor (Hamilton) limits
  - State 0 custom limits (Good / Okay)
  - State 3 custom limits (Good / Okay)

Domino:
    - Runs on port 8888 (Domino default for apps)
    - Binds to 0.0.0.0
    - Launched via app.sh

Local:
    python app.py
    Open http://localhost:8888
"""

import io
import json
import os
import re
from pathlib import Path

import pandas as pd
import numpy as np
import plotly
import plotly.graph_objects as go
from flask import Flask, render_template_string, request, jsonify

# ═════════════════════════════════════════════════════════════════════════
# Limit definitions
# ═════════════════════════════════════════════════════════════════════════

# Vendor (Hamilton) limits — symmetric
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

# State 0 custom limits — asymmetric (70% lower)
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

# State 3 custom limits — asymmetric (70% lower), post-e-conditioning
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

# Map of limit set key -> (label, limits_dict, color, dash_style)
LIMIT_SETS = {
    "vendor":      ("Vendor (Hamilton)",  VENDOR_LIMITS,       "red",       "dash"),
    "state0_good": ("State 0 — Good",    STATE0_LIMITS_GOOD,  "green",     "dashdot"),
    "state0_okay": ("State 0 — Okay",    STATE0_LIMITS_OKAY,  "orange",    "dot"),
    "state3_good": ("State 3 — Good",    STATE3_LIMITS_GOOD,  "#1b5e20",   "dashdot"),
    "state3_okay": ("State 3 — Okay",    STATE3_LIMITS_OKAY,  "#e65100",   "dot"),
}


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


# ── Plotly chart builders ────────────────────────────────────────────────

COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
]


def build_verification_figure(all_results, selected_limits, use_abs=False):
    """Build Plotly figure with data traces and selected limit bands."""
    fig = go.Figure()

    all_freqs = set()
    for _, fa in all_results:
        all_freqs.update(fa.keys())
    all_freqs = sorted(all_freqs)

    # ── Draw selected limit bands ──
    for limit_key in selected_limits:
        if limit_key not in LIMIT_SETS:
            continue
        label, limits_dict, color, dash = LIMIT_SETS[limit_key]
        limit_freqs = [f for f in all_freqs if f in limits_dict]
        if not limit_freqs:
            continue

        upper = [limits_dict[f][0] for f in limit_freqs]
        lower = [limits_dict[f][1] for f in limit_freqs]

        if use_abs:
            fig.add_trace(go.Scatter(
                x=limit_freqs, y=upper, mode="lines",
                line=dict(color=color, width=2, dash=dash),
                name=f"{label}",
            ))
        else:
            fig.add_trace(go.Scatter(
                x=limit_freqs, y=upper, mode="lines",
                line=dict(color=color, width=2, dash=dash),
                name=f"{label} (upper)",
            ))
            fig.add_trace(go.Scatter(
                x=limit_freqs, y=lower, mode="lines",
                line=dict(color=color, width=2, dash=dash),
                name=f"{label} (lower)",
                showlegend=False,
            ))

    # ── Draw data traces ──
    overall_pass = True
    for i, (meta, freq_avg) in enumerate(all_results):
        freqs = sorted(freq_avg.keys())
        values = [abs(freq_avg[f]) if use_abs else freq_avg[f] for f in freqs]
        color = COLORS[i % len(COLORS)]

        label = Path(meta.get("filename", "Unknown")).stem
        sensor = meta.get("sensor_serial", "")
        n = meta.get("n_measurements", "?")
        legend = f"{label} — S/N {sensor} (n={n})" if sensor else f"{label} (n={n})"

        hover = [f"<b>{label}</b><br>Freq: {f} kHz<br>Cap: {v:.4f} pF/cm"
                 for f, v in zip(freqs, values)]

        fig.add_trace(go.Scatter(
            x=freqs, y=values, mode="lines+markers",
            line=dict(color=color, width=2.5),
            marker=dict(size=7),
            name=legend,
            hovertext=hover, hoverinfo="text",
        ))

        # Check pass/fail against vendor limits
        for f in freqs:
            if f in VENDOR_LIMITS:
                u, l = VENDOR_LIMITS[f]
                if freq_avg[f] > u or freq_avg[f] < l:
                    overall_pass = False

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
            text="FScan Verification" + (" — Absolute Values" if use_abs else ""),
            font=dict(size=18),
        ),
        legend=dict(x=1.02, y=1, font=dict(size=10)),
        hovermode="closest",
        template="plotly_white",
        margin=dict(l=60, r=300, t=60, b=80),
        height=600,
    )

    status_text = "ALL PASS" if overall_pass else "FAIL DETECTED"
    status_color = "green" if overall_pass else "red"
    fig.add_annotation(
        x=0.99, y=0.99, xref="paper", yref="paper",
        text=f"<b>Vendor Status: {status_text}</b>",
        showarrow=False, font=dict(size=14, color="white"),
        bgcolor=status_color, borderpad=6, opacity=0.9,
        xanchor="right", yanchor="top",
    )

    return fig, overall_pass


def build_table_data(all_results):
    rows = []
    for meta, freq_avg in all_results:
        label = Path(meta.get("filename", "Unknown")).stem
        sensor = meta.get("sensor_serial", "")
        n = meta.get("n_measurements", "?")
        for freq in sorted(freq_avg.keys()):
            val = freq_avg[freq]
            # Vendor limits
            v_upper, v_lower = VENDOR_LIMITS.get(freq, (None, None))
            vendor_pass = (v_lower <= val <= v_upper) if v_upper is not None else None
            # State 0 Good
            s0g_upper, s0g_lower = STATE0_LIMITS_GOOD.get(freq, (None, None))
            s0g_pass = (s0g_lower <= val <= s0g_upper) if s0g_upper is not None else None
            # State 3 Good
            s3g_upper, s3g_lower = STATE3_LIMITS_GOOD.get(freq, (None, None))
            s3g_pass = (s3g_lower <= val <= s3g_upper) if s3g_upper is not None else None

            def status_str(passed):
                if passed is None:
                    return "N/A"
                return "PASS" if passed else "FAIL"

            rows.append({
                "file": label,
                "sensor": sensor,
                "n_scans": n,
                "freq": freq,
                "avg_cap": round(val, 4),
                "vendor": status_str(vendor_pass),
                "state0": status_str(s0g_pass),
                "state3": status_str(s3g_pass),
            })
    return rows


# ── Flask app ────────────────────────────────────────────────────────────

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB limit

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>FScan Verification</title>
    <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f7fa; color: #333;
        }
        header {
            background: linear-gradient(135deg, #1a237e, #283593);
            color: white; padding: 16px 32px;
            display: flex; align-items: center; gap: 16px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }
        header h1 { font-size: 22px; font-weight: 600; }
        header span { font-size: 13px; opacity: 0.8; }
        .container { max-width: 1400px; margin: 0 auto; padding: 24px; }
        .card {
            background: white; border-radius: 10px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.06);
            padding: 24px; margin-bottom: 20px;
        }
        .card h2 { font-size: 16px; color: #1a237e; margin-bottom: 16px; }
        .upload-zone {
            border: 2px dashed #90caf9; border-radius: 10px;
            padding: 40px; text-align: center; cursor: pointer;
            transition: all 0.2s; background: #f8fbff;
        }
        .upload-zone:hover, .upload-zone.dragging {
            border-color: #1a237e; background: #e8eaf6;
        }
        .upload-zone p { margin: 8px 0; color: #666; }
        .upload-zone .icon { font-size: 36px; margin-bottom: 8px; }
        .file-list {
            display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px;
        }
        .file-chip {
            display: inline-flex; align-items: center; gap: 6px;
            background: #e3f2fd; color: #1565c0; padding: 6px 12px;
            border-radius: 20px; font-size: 13px;
        }
        .file-chip .remove {
            cursor: pointer; color: #c62828; font-weight: bold;
            margin-left: 4px;
        }
        .controls {
            display: flex; gap: 12px; align-items: center;
            flex-wrap: wrap; margin-top: 16px;
        }
        button {
            padding: 10px 24px; border: none; border-radius: 6px;
            font-size: 14px; font-weight: 600; cursor: pointer;
            transition: all 0.15s;
        }
        .btn-primary {
            background: #1a237e; color: white;
        }
        .btn-primary:hover { background: #283593; }
        .btn-primary:disabled { background: #9e9e9e; cursor: default; }
        .btn-secondary {
            background: #e0e0e0; color: #333;
        }
        .btn-secondary:hover { background: #bdbdbd; }
        label.toggle {
            display: inline-flex; align-items: center; gap: 6px;
            font-size: 13px; cursor: pointer;
        }
        .limit-checks {
            display: flex; flex-wrap: wrap; gap: 12px; margin-top: 12px;
            padding: 12px; background: #fafafa; border-radius: 8px;
            border: 1px solid #e0e0e0;
        }
        .limit-checks label {
            display: inline-flex; align-items: center; gap: 4px;
            font-size: 13px; cursor: pointer;
        }
        .limit-checks .limit-color {
            display: inline-block; width: 20px; height: 3px;
            border-radius: 2px; vertical-align: middle;
        }
        .spinner {
            display: inline-block; width: 20px; height: 20px;
            border: 3px solid #e0e0e0; border-top: 3px solid #1a237e;
            border-radius: 50%; animation: spin 0.8s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        #status-msg {
            font-size: 13px; color: #666; margin-left: 12px;
        }
        .plot-container { min-height: 400px; }
        .meta-grid {
            display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 12px; margin-bottom: 16px;
        }
        .meta-item {
            background: #fafafa; border: 1px solid #e0e0e0;
            border-radius: 8px; padding: 12px;
        }
        .meta-item h3 { font-size: 14px; color: #1a237e; margin-bottom: 6px; }
        .meta-item p { font-size: 13px; color: #555; margin: 2px 0; }
        .table-wrap { overflow-x: auto; max-height: 500px; overflow-y: auto; }
        table {
            width: 100%; border-collapse: collapse; font-size: 13px;
        }
        th {
            position: sticky; top: 0; background: #1a237e; color: white;
            padding: 10px 12px; text-align: left; font-weight: 600;
        }
        td { padding: 8px 12px; border-bottom: 1px solid #eee; }
        tr:hover td { background: #f5f5f5; }
        .status-pass { color: #2e7d32; font-weight: 600; }
        .status-fail { color: #c62828; font-weight: 600; background: #ffebee; }
        .tab-bar {
            display: flex; gap: 0; border-bottom: 2px solid #e0e0e0;
            margin-bottom: 16px;
        }
        .tab-btn {
            padding: 10px 20px; background: none; border: none;
            border-bottom: 3px solid transparent;
            font-size: 14px; font-weight: 600; color: #666;
            cursor: pointer; transition: all 0.15s;
        }
        .tab-btn.active { color: #1a237e; border-bottom-color: #1a237e; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .domino-note {
            font-size: 11px; color: #999; text-align: center;
            padding: 12px; border-top: 1px solid #eee; margin-top: 20px;
        }
    </style>
</head>
<body>
    <header>
        <div>
            <h1>FScan Probe Verification</h1>
            <span>Capacitance probe fscan verification — Vendor + State 0 + State 3 limits</span>
        </div>
    </header>

    <div class="container">
        <div class="card">
            <h2>Upload FScan Excel Files</h2>
            <div class="upload-zone" id="dropZone">
                <div class="icon">&#x1F4C2;</div>
                <p><strong>Drag &amp; drop</strong> Excel files here or <strong>click to browse</strong></p>
                <p style="font-size:12px; color:#999;">Accepts .xlsx / .xls &mdash; multiple files for overlay</p>
            </div>
            <input type="file" id="fileInput" multiple accept=".xlsx,.xls" style="display:none">
            <div class="file-list" id="fileList"></div>

            <h2 style="margin-top:20px;">Select Limit Sets</h2>
            <div class="limit-checks">
                <label>
                    <input type="checkbox" name="limits" value="vendor" checked>
                    <span class="limit-color" style="background:red;"></span>
                    Vendor (Hamilton)
                </label>
                <label>
                    <input type="checkbox" name="limits" value="state0_good" checked>
                    <span class="limit-color" style="background:green;"></span>
                    State 0 — Good
                </label>
                <label>
                    <input type="checkbox" name="limits" value="state0_okay">
                    <span class="limit-color" style="background:orange;"></span>
                    State 0 — Okay
                </label>
                <label>
                    <input type="checkbox" name="limits" value="state3_good" checked>
                    <span class="limit-color" style="background:#1b5e20;"></span>
                    State 3 — Good
                </label>
                <label>
                    <input type="checkbox" name="limits" value="state3_okay">
                    <span class="limit-color" style="background:#e65100;"></span>
                    State 3 — Okay
                </label>
            </div>

            <div class="controls">
                <button class="btn-primary" id="genBtn" disabled>Generate Plot</button>
                <button class="btn-secondary" id="clearBtn">Clear All</button>
                <label class="toggle">
                    <input type="checkbox" id="absToggle"> Absolute values
                </label>
                <span id="status-msg"></span>
            </div>
        </div>

        <div id="results" style="display:none;">
            <div class="card">
                <h2>File Summary</h2>
                <div class="meta-grid" id="metaGrid"></div>
            </div>

            <div class="card">
                <div class="tab-bar">
                    <button class="tab-btn active" data-tab="plotTab">&#x1F4C8; Verification Plot</button>
                    <button class="tab-btn" data-tab="tableTab">&#x1F4CB; Data Table</button>
                </div>
                <div id="plotTab" class="tab-content active">
                    <div class="plot-container" id="plotDiv"></div>
                </div>
                <div id="tableTab" class="tab-content">
                    <div class="table-wrap" id="tableWrap"></div>
                </div>
            </div>
        </div>

        <div class="domino-note">
            FScan Verification Tool &mdash; Capacitance Probe Analysis
        </div>
    </div>

    <script>
        const dropZone = document.getElementById("dropZone");
        const fileInput = document.getElementById("fileInput");
        const fileList = document.getElementById("fileList");
        const genBtn = document.getElementById("genBtn");
        const clearBtn = document.getElementById("clearBtn");
        const absToggle = document.getElementById("absToggle");
        const statusMsg = document.getElementById("status-msg");
        const results = document.getElementById("results");

        let uploadedFiles = [];

        dropZone.addEventListener("click", () => fileInput.click());
        dropZone.addEventListener("dragover", e => { e.preventDefault(); dropZone.classList.add("dragging"); });
        dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));
        dropZone.addEventListener("drop", e => {
            e.preventDefault(); dropZone.classList.remove("dragging");
            addFiles(e.dataTransfer.files);
        });
        fileInput.addEventListener("change", () => { addFiles(fileInput.files); fileInput.value = ""; });

        function addFiles(files) {
            for (const f of files) {
                if (!uploadedFiles.some(u => u.name === f.name)) {
                    uploadedFiles.push(f);
                }
            }
            renderFileList();
        }

        function renderFileList() {
            fileList.innerHTML = "";
            uploadedFiles.forEach((f, i) => {
                const chip = document.createElement("span");
                chip.className = "file-chip";
                chip.textContent = f.name + " ";
                const rm = document.createElement("span");
                rm.className = "remove";
                rm.textContent = "\u2715";
                rm.dataset.idx = i;
                rm.addEventListener("click", e => {
                    uploadedFiles.splice(+e.target.dataset.idx, 1);
                    renderFileList();
                });
                chip.appendChild(rm);
                fileList.appendChild(chip);
            });
            genBtn.disabled = uploadedFiles.length === 0;
        }

        clearBtn.addEventListener("click", () => {
            uploadedFiles = [];
            renderFileList();
            results.style.display = "none";
            statusMsg.textContent = "";
        });

        function getSelectedLimits() {
            return Array.from(document.querySelectorAll('input[name="limits"]:checked'))
                        .map(cb => cb.value);
        }

        genBtn.addEventListener("click", async () => {
            if (uploadedFiles.length === 0) return;
            genBtn.disabled = true;
            statusMsg.innerHTML = '<span class="spinner"></span> Processing\u2026';

            const fd = new FormData();
            uploadedFiles.forEach(f => fd.append("files", f));
            fd.append("use_abs", absToggle.checked ? "1" : "0");
            getSelectedLimits().forEach(l => fd.append("limits", l));

            try {
                const resp = await fetch("analyze", { method: "POST", body: fd });
                if (!resp.ok) {
                    const err = await resp.json();
                    throw new Error(err.error || "Server error");
                }
                const data = await resp.json();
                renderResults(data);
                statusMsg.textContent = "Done \u2014 " + data.file_count + " file(s) processed.";
            } catch (e) {
                statusMsg.textContent = "Error: " + e.message;
            } finally {
                genBtn.disabled = false;
            }
        });

        function renderResults(data) {
            results.style.display = "block";
            const metaGrid = document.getElementById("metaGrid");
            metaGrid.innerHTML = "";
            data.metadata.forEach(m => {
                const div = document.createElement("div");
                div.className = "meta-item";
                div.innerHTML = "<h3>" + (m.filename || "Unknown") + "</h3>"
                    + "<p><b>Sensor S/N:</b> " + (m.sensor_serial || "N/A") + "</p>"
                    + "<p><b>Batch:</b> " + (m.batch_name || "N/A") + "</p>"
                    + "<p><b>Date:</b> " + (m.creation_date || "N/A") + "</p>"
                    + "<p><b>Valid Scans:</b> " + (m.n_measurements || "?") + "</p>";
                metaGrid.appendChild(div);
            });

            const plotData = JSON.parse(data.plot_json);
            Plotly.newPlot("plotDiv", plotData.data, plotData.layout, {responsive: true});

            renderTable(data.table);
            results.scrollIntoView({ behavior: "smooth" });
        }

        function renderTable(rows) {
            const wrap = document.getElementById("tableWrap");
            if (!rows || rows.length === 0) { wrap.innerHTML = "<p>No data.</p>"; return; }
            let html = "<table><thead><tr>"
                + "<th>File</th><th>Sensor</th><th>Scans</th>"
                + "<th>Freq (kHz)</th><th>Avg Cap (pF/cm)</th>"
                + "<th>Vendor</th><th>State 0</th><th>State 3</th>"
                + "</tr></thead><tbody>";
            rows.forEach(r => {
                function cls(s) { return s === "PASS" ? "status-pass" : (s === "FAIL" ? "status-fail" : ""); }
                html += "<tr>"
                    + "<td>" + r.file + "</td><td>" + r.sensor + "</td><td>" + r.n_scans + "</td>"
                    + "<td>" + r.freq + "</td><td>" + r.avg_cap + "</td>"
                    + "<td class='" + cls(r.vendor) + "'>" + r.vendor + "</td>"
                    + "<td class='" + cls(r.state0) + "'>" + r.state0 + "</td>"
                    + "<td class='" + cls(r.state3) + "'>" + r.state3 + "</td>"
                    + "</tr>";
            });
            html += "</tbody></table>";
            wrap.innerHTML = html;
        }

        document.querySelectorAll(".tab-btn").forEach(btn => {
            btn.addEventListener("click", () => {
                document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
                document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));
                btn.classList.add("active");
                document.getElementById(btn.dataset.tab).classList.add("active");
                if (btn.dataset.tab === "plotTab") {
                    Plotly.Plots.resize(document.getElementById("plotDiv"));
                }
            });
        });
    </script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/analyze", methods=["POST"])
def analyze():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    use_abs = request.form.get("use_abs", "0") == "1"
    selected_limits = request.form.getlist("limits")
    if not selected_limits:
        selected_limits = ["vendor", "state0_good", "state3_good"]

    all_results = []
    all_metadata = []
    errors = []

    for f in files:
        if not f.filename:
            continue
        ext = Path(f.filename).suffix.lower()
        if ext not in (".xlsx", ".xls"):
            errors.append(f"{f.filename}: invalid file type")
            continue
        try:
            file_bytes = f.read()
            meta, freq_avg = parse_fscan_bytes(file_bytes, f.filename)
            all_results.append((meta, freq_avg))
            all_metadata.append(meta)
        except Exception as e:
            errors.append(f"{f.filename}: {str(e)}")

    if not all_results:
        return jsonify({"error": "No valid files. " + "; ".join(errors)}), 400

    fig, overall_pass = build_verification_figure(all_results, selected_limits, use_abs)
    table = build_table_data(all_results)

    return jsonify({
        "file_count": len(all_results),
        "overall_pass": overall_pass,
        "metadata": all_metadata,
        "plot_json": json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder),
        "table": table,
        "errors": errors,
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8888))
    app.run(host="0.0.0.0", port=port, debug=False)
