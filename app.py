
from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
import streamlit as st
from sqlalchemy import select

from db import init_db, get_session, Project

RAG_CHOICES = ["Green", "Amber", "Red"]


def projects_to_df(projects):
    rows = []
    for p in projects:
        rows.append({
            "ID": p.id,
            "Client Name": p.client_name,
            "Project Owner": p.project_owner,
            "Project Name": p.project_name,
            "Budget / PO Name": p.budget_po_name,
            "Schedule / Timeline": p.schedule_timeline,
            "Budget / Cost": p.rag_budget_cost,
            "Scope / Requirement": p.rag_scope_requirement,
            "Quality": p.rag_quality,
            "Risk": p.rag_risk,
            "Resource / Team": p.rag_resource_team,
            "MOHI": p.rag_mohi,
            "Notes": p.notes,
            "Updated By": p.updated_by,
            "Updated At (UTC)": p.updated_at,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame([{
            "ID": None,
            "Client Name": "",
            "Project Owner": "",
            "Project Name": "",
            "Budget / PO Name": "",
            "Schedule / Timeline": "",
            "Budget / Cost": "Green",
            "Scope / Requirement": "Green",
            "Quality": "Green",
            "Risk": "Green",
            "Resource / Team": "Green",
            "MOHI": "",
            "Notes": "",
            "Updated By": "",
            "Updated At (UTC)": "",
        }])
    return df


def upsert_from_df(df: pd.DataFrame, updated_by: str | None = None):
    session = get_session()
    try:
        for _, row in df.iterrows():
            # Skip blank rows
            if not str(row.get("Client Name", "")).strip() and not str(row.get("Project Name", "")).strip():
                continue

            rid = row.get("ID")
            obj = None
            if pd.notna(rid):
                obj = session.get(Project, int(rid))

            if obj is None:
                obj = Project()
                session.add(obj)

            obj.client_name = str(row.get("Client Name", "")).strip()
            obj.project_owner = str(row.get("Project Owner", "")).strip()
            obj.project_name = str(row.get("Project Name", "")).strip()
            obj.budget_po_name = (str(row.get("Budget / PO Name", "")).strip() or None)
            obj.schedule_timeline = (str(row.get("Schedule / Timeline", "")).strip() or None)

            obj.rag_budget_cost = (str(row.get("Budget / Cost", "Green")).strip() or "Green")
            obj.rag_scope_requirement = (str(row.get("Scope / Requirement", "Green")).strip() or "Green")
            obj.rag_quality = (str(row.get("Quality", "Green")).strip() or "Green")
            obj.rag_risk = (str(row.get("Risk", "Green")).strip() or "Green")
            obj.rag_resource_team = (str(row.get("Resource / Team", "Green")).strip() or "Green")
            obj.rag_mohi = (str(row.get("MOHI", "")).strip() or None)

            obj.notes = (str(row.get("Notes", "")).strip() or None)
            obj.updated_by = (updated_by or str(row.get("Updated By", "")).strip() or None)
            obj.updated_at = datetime.utcnow()

            # Validate RAG values
            for col_name, val in [
                ("Budget / Cost", obj.rag_budget_cost),
                ("Scope / Requirement", obj.rag_scope_requirement),
                ("Quality", obj.rag_quality),
                ("Risk", obj.rag_risk),
                ("Resource / Team", obj.rag_resource_team),
            ]:
                if val not in RAG_CHOICES:
                    raise ValueError(f"Invalid RAG value '{val}' in column {col_name} (must be Green/Amber/Red)")

        session.commit()
    finally:
        session.close()


def delete_ids(ids):
    session = get_session()
    try:
        for rid in ids:
            obj = session.get(Project, int(rid))
            if obj:
                session.delete(obj)
        session.commit()
    finally:
        session.close()


def fetch_projects(filters: dict):
    session = get_session()
    try:
        stmt = select(Project)
        if filters.get("client"):
            stmt = stmt.where(Project.client_name.contains(filters["client"]))
        if filters.get("owner"):
            stmt = stmt.where(Project.project_owner.contains(filters["owner"]))
        if filters.get("project"):
            stmt = stmt.where(Project.project_name.contains(filters["project"]))
        if filters.get("any_rag") and filters["any_rag"] != "All":
            v = filters["any_rag"]
            stmt = stmt.where(
                (Project.rag_budget_cost == v) |
                (Project.rag_scope_requirement == v) |
                (Project.rag_quality == v) |
                (Project.rag_risk == v) |
                (Project.rag_resource_team == v)
            )
        stmt = stmt.order_by(Project.client_name, Project.project_name)
        return list(session.scalars(stmt).all())
    finally:
        session.close()


st.set_page_config(page_title="RAG Report Web App", layout="wide")
init_db()

st.title("RAG Report Web App")
st.caption("Excel-like entry for project health (Green / Amber / Red). Data stored locally in SQLite.")

with st.sidebar:
    st.header("Filters")
    client = st.text_input("Client contains")
    owner = st.text_input("Owner contains")
    project = st.text_input("Project contains")
    any_rag = st.selectbox("Any RAG equals", ["All"] + RAG_CHOICES)
    st.divider()
    updated_by = st.text_input("Your name (for audit)", value="")

    st.header("Import / Export")
    uploaded = st.file_uploader(
        "Import from Excel (.xlsx)",
        type=["xlsx"],
        help="Imports a sheet named 'RAG Report' if present; otherwise uses the first sheet."
    )
    export_btn = st.button("Generate Export")

# Import
if uploaded is not None:
    try:
        xls = pd.ExcelFile(uploaded)
        sheet = "RAG Report" if "RAG Report" in xls.sheet_names else xls.sheet_names[0]
        imp = pd.read_excel(xls, sheet_name=sheet)

        # Ensure required columns exist (best-effort)
        for col, default in [
            ("Client Name", ""),
            ("Project Owner", ""),
            ("Project Name", ""),
            ("Budget / PO Name", ""),
            ("Schedule / Timeline", ""),
            ("Budget / Cost", "Green"),
            ("Scope / Requirement", "Green"),
            ("Quality", "Green"),
            ("Risk", "Green"),
            ("Resource / Team", "Green"),
            ("MOHI", ""),
            ("Notes", ""),
        ]:
            if col not in imp.columns:
                imp[col] = default

        keep = [
            "Client Name","Project Owner","Project Name","Budget / PO Name","Schedule / Timeline",
            "Budget / Cost","Scope / Requirement","Quality","Risk","Resource / Team","MOHI","Notes"
        ]
        imp = imp[keep]

        upsert_from_df(imp, updated_by=updated_by.strip() or None)
        st.success(f"Imported {len(imp)} rows from '{sheet}'")
    except Exception as e:
        st.error(f"Import failed: {e}")

# Fetch and show grid
projects = fetch_projects({
    "client": client.strip(),
    "owner": owner.strip(),
    "project": project.strip(),
    "any_rag": any_rag
})

st.subheader("Projects")

edited = st.data_editor(
    projects_to_df(projects),
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "ID": st.column_config.NumberColumn(disabled=True),
        "Budget / Cost": st.column_config.SelectboxColumn(options=RAG_CHOICES, required=True),
        "Scope / Requirement": st.column_config.SelectboxColumn(options=RAG_CHOICES, required=True),
        "Quality": st.column_config.SelectboxColumn(options=RAG_CHOICES, required=True),
        "Risk": st.column_config.SelectboxColumn(options=RAG_CHOICES, required=True),
        "Resource / Team": st.column_config.SelectboxColumn(options=RAG_CHOICES, required=True),
        "Updated By": st.column_config.TextColumn(disabled=True),
        "Updated At (UTC)": st.column_config.TextColumn(disabled=True),
    },
    hide_index=True,
)

col1, col2, col3 = st.columns([1, 1, 3])
with col1:
    if st.button("Save changes", type="primary"):
        try:
            upsert_from_df(edited, updated_by=updated_by.strip() or None)
            st.success("Saved")
            st.rerun()
        except Exception as e:
            st.error(f"Save failed: {e}")

with col2:
    delete_text = st.text_input("Delete IDs (comma-separated)", value="", help="Example: 1,2,3")
    if st.button("Delete"):
        try:
            ids = [int(x.strip()) for x in delete_text.split(',') if x.strip()]
            delete_ids(ids)
            st.success(f"Deleted {len(ids)}")
            st.rerun()
        except Exception as e:
            st.error(f"Delete failed: {e}")

# Export
if export_btn:
    export_projects = fetch_projects({"client": "", "owner": "", "project": "", "any_rag": "All"})
    export_df = projects_to_df(export_projects).drop(columns=["ID"], errors='ignore')

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="RAG Report", index=False)
        pd.DataFrame({"RAG": RAG_CHOICES}).to_excel(writer, sheet_name="Meta", index=False)
    out.seek(0)

    st.download_button(
        label="Download RAG Report.xlsx",
        data=out,
        file_name="RAG Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()
st.subheader("Quick stats")

if projects:
    def non_green(p: Project) -> bool:
        vals = [p.rag_budget_cost, p.rag_scope_requirement, p.rag_quality, p.rag_risk, p.rag_resource_team]
        return any(v != 'Green' for v in vals)

    total = len(projects)
    amber = sum(1 for p in projects if 'Amber' in [p.rag_budget_cost, p.rag_scope_requirement, p.rag_quality, p.rag_risk, p.rag_resource_team])
    red = sum(1 for p in projects if 'Red' in [p.rag_budget_cost, p.rag_scope_requirement, p.rag_quality, p.rag_risk, p.rag_resource_team])
    all_green = sum(1 for p in projects if not non_green(p))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Projects (filtered)", total)
    c2.metric("All Green", all_green)
    c3.metric("Has Amber", amber)
    c4.metric("Has Red", red)
else:
    st.info("No rows yet. Add a row in the grid and click 'Save changes'.")

with st.expander("Admin / Notes"):
    st.markdown(
        "- Storage: SQLite file `rag_report.db` (configurable via env var `RAG_DB_PATH`).
"
        "- RAG values are validated to be one of **Green / Amber / Red**.
"
        "- Import expects a sheet named `RAG Report` if present; otherwise it uses the first sheet."
    )
