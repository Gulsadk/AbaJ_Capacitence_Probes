#!/bin/bash
# Domino App launcher — Root dispatcher
#
# Set the DOMINO_APP environment variable to choose which app to run:
#   csv_to_excel       — CSV/Excel log data → model-ready Excel converter
#   fscan_verification — FScan verification plot (Flask)
#   overlay            — Probe overlay over time (Streamlit)
#
# Default: csv_to_excel

APP_NAME="${DOMINO_APP:-csv_to_excel}"

echo "=== Starting app: ${APP_NAME} ==="

cd "${APP_NAME}" || { echo "ERROR: App folder '${APP_NAME}' not found"; exit 1; }

pip install -r requirements.txt --quiet --force-reinstall --no-deps
pip install -r requirements.txt --quiet

if [ "$APP_NAME" = "fscan_verification" ]; then
    python app.py
else
    streamlit run app.py \
        --server.port 8888 \
        --server.address 0.0.0.0 \
        --server.headless true \
        --browser.gatherUsageStats false
fi
