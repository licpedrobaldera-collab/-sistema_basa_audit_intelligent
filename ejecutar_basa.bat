@echo off
TITLE BASA - Audit Intelligent USA - V4
COLOR 0A
mkdir evidencias_auditoria 2>nul
mkdir reportes_exportados 2>nul
pip install streamlit pandas openpyxl python-docx PyMuPDF plotly pillow >nul
python -m streamlit run sistema_basa_audit_intelligent.py --server.address=0.0.0.0 --server.port=8501
pause
