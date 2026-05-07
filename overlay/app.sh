#!/bin/bash
# Domino App launcher — Probe Overlay (Streamlit)
# Domino expects the app to listen on port 8888

pip install -r requirements.txt --quiet

streamlit run app.py \
    --server.port 8888 \
    --server.address 0.0.0.0 \
    --server.headless true \
    --browser.gatherUsageStats false
