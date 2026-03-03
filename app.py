
from __future__ import annotations

import io
import os
import sqlite3
from datetime import datetime

import pandas as pd
import streamlit as st

# ----------------------------
# Configuration
# ----------------------------
RAG_CHOICES = ["Green", "Amber", "Red"]
DB_PATH = os.getenv("RAG_DB_PATH", "rag_report.db")

COLS = [
    "ID",
    "Client Name",
    "Project Owner",
    "Project Name",
    "Budget / PO Name",
    "Schedule / Timeline",
    "Budget / Cost",
    "Scope / Requirement",
    "Quality",
    "Risk",
    "Resource / Team",
    "MOHI",
    "Notes",
    "Updated By",
    "Updated At (UTC)",
]


def get_conn():
    return sqlite3.connect(DB_PATH)


def init_db():
    with get_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_name TEXT NOT NULL,
                project_owner TEXT NOT NULL,
                project_name TEXT NOT NULL,
                budget_po_name TEXT,
                schedule_timeline TEXT,
                rag_budget_cost TEXT NOT NULL DEFAULT 'Green',
                rag_scope_requirement TEXT NOT NULL DEFAULT 'Green',
                rag_quality TEXT NOT NULL DEFAULT 'Green',
                rag_risk TEXT NOT NULL DEFAULT 'Green',
                rag_resource_team TEXT NOT NULL DEFAULT 'Green',
                rag_mohi TEXT,
                notes TEXT,
                updated_by TEXT,
                updated_at_utc TEXT NOT NULL
            );
            """
        )


def fetch_projects(filters: dict) -> pd.DataFrame:
    q = "SELECT * FROM projects WHERE 1=1"
    params = []

    if filters.get("client"):
        q += " AND client_name LIKE ?"
        params.append(f"%{filters['client']}%")
    if filters.get("owner"):
        q += " AND project_owner LIKE ?"
        params.append(f"%{filters['owner']}%")
    if filters.get("project"):
        q += " AND project_name LIKE ?"
        params.append(f"%{filters['project']}%")

    any_rag = filters.get("any_rag")
    if any_rag and any_rag != "All":
        q += " AND (rag_budget_cost = ? OR rag_scope_requirement = ? OR rag_quality = ? OR rag_risk = ? OR rag_resource_team = ?)"
        params.extend([any_rag] * 5)

    q += " ORDER BY client_name, project_name"

    with get_conn() as conn:
        df = pd.read_sql_query(q, conn, params=params)

    if df.empty:
        return pd.DataFrame(columns=COLS)

    df = df.rename(
        columns={
            "id": "ID",
            "client_name": "Client Name",
            "project_owner": "Project Owner",
            "project_name": "Project Name",
            "budget_po_name": "Budget / PO Name",
            "schedule_timeline": "Schedule / Timeline",
            "rag_budget_cost": "Budget / Cost",
            "rag_scope_requirement": "Scope / Requirement",
            "rag_quality": "Quality",
            "rag_risk": "Risk",
            "rag_resource_team": "Resource / Team",
            "rag_mohi": "MOHI",
            "notes": "Notes",
            "updated_by": "Updated By",
            "updated_at_utc": "Updated At (UTC)",
        }
    )

    for c in COLS:
        if c not in df.columns:
            df[c] = ""

    return df[COLS]


def validate_rag_values(df: pd.DataFrame):
    for col in ["Budget / Cost", "Scope / Requirement", "Quality", "Risk", "Resource / Team"]:
        bad = df[~df[col].isin(RAG_CHOICES) & df[col].notna()]
        if not bad.empty:
            vals = bad[col].astype(str).unique().tolist()
            raise ValueError(f"Invalid RAG value(s) in '{col}': {vals}. Allowed: {RAG_CHOICES}")


def upsert_from_df(df: pd.DataFrame, updated_by: str | None):
    for col, default in [
        ("Budget / Cost", "Green"),
        ("Scope / Requirement", "Green"),
        ("Quality", "Green"),
        ("Risk", "Green"),
        ("Resource / Team", "Green"),
    ]:
        if col in df.columns:
            df[col] = df[col].fillna(default).replace("", default)

    validate_rag_values(df)

    now_utc = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

    with get_conn() as conn:
        cur = conn.cursor()
        for _, row in df.iterrows():
            if (str(row.get("Client Name", "")).strip() == "") and (str(row.get("Project Name", "")).strip() == ""):
                continue

            client = str(row.get("Client Name", "")).strip()
            owner = str(row.get("Project Owner", "")).strip()
            pname = str(row.get("Project Name", "")).strip()

            if not client or not owner or not pname:
                continue

            values = (
                client,
                owner,
                pname,
                (str(row.get("Budget / PO Name", "")).strip() or None),
                (str(row.get("Schedule / Timeline", "")).strip() or None),
                str(row.get("Budget / Cost", "Green")).strip(),
                str(row.get("Scope / Requirement", "Green")).strip(),
                str(row.get("Quality", "Green")).strip(),
                str(row.get("Risk", "Green")).strip(),
                str(row.get("Resource / Team", "Green")).strip(),
                (str(row.get("MOHI", "")).strip() or None),
                (str(row.get("Notes", "")).strip() or None),
                (updated_by or str(row.get("Updated By", "")).strip() or None),
                now_utc,
            )

            rid = row.get("ID")
            if pd.notna(rid) and str(rid).strip() != "":
                cur.execute(
                    """
                    UPDATE projects
                    SET client_name=?, project_owner=?, project_name=?, budget_po_name=?, schedule_timeline=?,
                        rag_budget_cost=?, rag_scope_requirement=?, rag_quality=?, rag_risk=?, rag_resource_team=?,
                        rag_mohi=?, notes=?, updated_by=?, updated_at_utc=?
                    WHERE id=?
                    """,
                    (*values, int(rid)),
                )
            else:
                cur.execute(
                    """
                    INSERT INTO projects (
                        client_name, project_owner, project_name, budget_po_name, schedule_timeline,
                        rag_budget_cost, rag_scope_requirement, rag_quality, rag_risk, rag_resource_team,
                        rag_mohi, notes, updated_by, updated_at_utc
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    values,
                )

        conn.commit()


def delete_ids(ids: list[int]):
    if not ids:
        return
    with get_conn() as conn:
        conn.executemany("DELETE FROM projects WHERE id=?", [(i,) for i in ids])
        conn.commit()


# ----------------------------
# UI
# ----------------------------

st.set_page_config(page_title="RAG Report Web App", layout="wide")
init_db()

st.title("RAG Report Web App")
st.caption("Excel-like entry for project health (Green / Amber / Red).")

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

if uploaded is not None:
    try:
        xls = pd.ExcelFile(uploaded)
        sheet = "RAG Report" if "RAG Report" in xls.sheet_names else xls.sheet_names[0]
        imp = pd.read_excel(xls, sheet_name=sheet)

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
        st.success(f"Imported {len(imp)} row(s) from '{sheet}'")
    except Exception as e:
        st.error(f"Import failed: {e}")

current = fetch_projects({
    "client": client.strip(),
    "owner": owner.strip(),
    "project": project.strip(),
    "any_rag": any_rag,
})

if current.empty:
    current = pd.DataFrame([{
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

st.subheader("Projects")

edited = st.data_editor(
    current,
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

c1, c2, _ = st.columns([1, 1, 3])

with c1:
    if st.button("Save changes", type="primary"):
        try:
            upsert_from_df(edited, updated_by=updated_by.strip() or None)
            st.success("Saved")
            st.rerun()
        except Exception as e:
            st.error(f"Save failed: {e}")

with c2:
    delete_text = st.text_input("Delete IDs (comma-separated)", value="", help="Example: 1,2,3")
    if st.button("Delete"):
        try:
            ids = [int(x.strip()) for x in delete_text.split(",") if x.strip()]
            delete_ids(ids)
            st.success(f"Deleted {len(ids)}")
            st.rerun()
        except Exception as e:
            st.error(f"Delete failed: {e}")

if export_btn:
    all_df = fetch_projects({"client": "", "owner": "", "project": "", "any_rag": "All"})
    export_df = all_df.drop(columns=["ID"], errors="ignore")

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

live = fetch_projects({"client": client.strip(), "owner": owner.strip(), "project": project.strip(), "any_rag": any_rag})

if not live.empty:
    amber = (live[["Budget / Cost","Scope / Requirement","Quality","Risk","Resource / Team"]] == "Amber").any(axis=1).sum()
    red = (live[["Budget / Cost","Scope / Requirement","Quality","Risk","Resource / Team"]] == "Red").any(axis=1).sum()
    all_green = (live[["Budget / Cost","Scope / Requirement","Quality","Risk","Resource / Team"]] == "Green").all(axis=1).sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Projects (filtered)", int(len(live)))
    k2.metric("All Green", int(all_green))
    k3.metric("Has Amber", int(amber))
    k4.metric("Has Red", int(red))
else:
    st.info("No rows yet. Add a row in the grid and click 'Save changes'.")

with st.expander("Admin / Notes"):
    st.markdown(
        """- Storage: SQLite file `rag_report.db` (configurable via env var `RAG_DB_PATH`).
- RAG values are validated to be one of **Green / Amber / Red**.
- Import expects a sheet named `RAG Report` if present; otherwise it uses the first sheet."""
    )
