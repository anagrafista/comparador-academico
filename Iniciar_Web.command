#!/bin/bash
# Lanzador automático de la versión Web (Streamlit) para macOS
cd "$(dirname "$0")"
python3 -m streamlit run streamlit_app.py
