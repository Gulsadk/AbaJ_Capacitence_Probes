#!/bin/bash
# Domino App launcher — Unified Capacitance Probe Tools
#
# Single Streamlit app with landing page + 3 tools:
#   1. CSV → Excel converter
#   2. FScan verification plot
#   3. Probe overlay over time

echo "=== Installing requirements ==="
pip install -r requirements.txt --quiet

echo "=== Starting Capacitance Probe Tools ==="
streamlit run app.py \
    --server.port 8888 \
    --server.address 0.0.0.0 \
    --server.headless true \
    --browser.gatherUsageStats false
