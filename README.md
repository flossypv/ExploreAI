
# RAG Report Web App (Streamlit)

A simple Excel-like web app to enter and maintain project RAG status.

## Features
- Excel-like grid entry with dropdowns for RAG fields (Green/Amber/Red)
- Filters (client/owner/project and Any-RAG)
- Import from Excel (.xlsx)
- Export to Excel (creates sheets: `RAG Report` and `Meta`)
- Local persistence using SQLite

## Run locally
```bash
cd rag_streamlit_app
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Configuration
- Database path: set environment variable `RAG_DB_PATH` (default: `rag_report.db`).

## Notes
- MVP codebase.
- For multi-user concurrency or enterprise auth, deploy behind your org’s SSO gateway or move persistence to a managed DB.
