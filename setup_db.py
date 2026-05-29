"""
Smart Material Estimator — setup_db.py
=======================================
One-time script: converts your three Excel files into a normalised
SQLite database (sme_database.db).

Run ONCE from the project folder:
    python setup_db.py

Re-run any time you update the source Excel files to refresh master data.
Consumption logs and SQM progress are preserved across re-runs.
"""

import os, sys, sqlite3
import pandas as pd
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "sme_database.db")
PATH_A   = os.path.join(BASE_DIR, "Materials_DetailsAvailable_Qty.xlsx")
PATH_B   = os.path.join(BASE_DIR, "For_1_SQM.xlsx")
PATH_C   = os.path.join(BASE_DIR, "Equipment.xlsx")
SHEET_A, SHEET_B, SHEET_C = "Materials", "LINING SYSTEM MATERIAL CONSM", "Data Input"

# ─────────────────────────────────────────────────────────────────────────────
# DATABASE SCHEMA
# ─────────────────────────────────────────────────────────────────────────────
SCHEMA = """
-- ── Master: inventory (one row per unique Material_Code) ───────────────────
CREATE TABLE IF NOT EXISTS inventory (
    material_code    TEXT    PRIMARY KEY,
    material_name    TEXT,
    nature           TEXT,
    uom              TEXT,
    available_qty    REAL    NOT NULL DEFAULT 0,
    ordered_qty      REAL    NOT NULL DEFAULT 0
);

-- ── Master: lining system recipe (For_1_SQM per material per system code) ──
CREATE TABLE IF NOT EXISTS recipe (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    lining_system_code       TEXT NOT NULL,
    lining_system_short_name TEXT,
    lining_type              TEXT,
    material_code            TEXT NOT NULL,
    material_description     TEXT,
    material_name            TEXT,
    for_1_sqm                REAL,
    uom                      TEXT,
    FOREIGN KEY (material_code) REFERENCES inventory(material_code)
);

-- ── Master: equipment surface areas (one row per tag+code+area section) ────
CREATE TABLE IF NOT EXISTS equipment (
    id                       INTEGER PRIMARY KEY AUTOINCREMENT,
    location                 TEXT,
    type                     TEXT,
    lining_system_code       TEXT,
    lining_system_short_name TEXT,
    lining_type              TEXT,
    equipment_tag            TEXT NOT NULL,
    name                     TEXT,
    description              TEXT,
    material_spec            TEXT,
    design                   TEXT,
    surface_area_sqm         REAL,
    lining_systems           TEXT
);

-- ── Live: SQM progress (tracks completed SQM per tag+code) ─────────────────
-- done_sqm starts at 0 and increases with each consumption entry.
-- remaining_sqm = original_sqm - done_sqm
CREATE TABLE IF NOT EXISTS sqm_progress (
    equipment_tag      TEXT NOT NULL,
    lining_system_code TEXT NOT NULL,
    original_sqm       REAL NOT NULL DEFAULT 0,
    done_sqm           REAL NOT NULL DEFAULT 0,
    PRIMARY KEY (equipment_tag, lining_system_code)
);

-- ── Live: daily consumption log ────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS consumption_log (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date           TEXT    NOT NULL,
    equipment_tag        TEXT    NOT NULL,
    lining_system_code   TEXT    NOT NULL,
    lining_system_name   TEXT,
    sqm_completed        REAL    NOT NULL DEFAULT 0,
    material_code        TEXT    NOT NULL,
    material_name        TEXT,
    uom                  TEXT,
    expected_qty         REAL,
    consumed_qty         REAL    NOT NULL DEFAULT 0,
    notes                TEXT,
    submitted_at         TEXT    DEFAULT (datetime('now'))
);

-- ── Live: material receipt log ────────────────────────────────────────────
-- Every receipt entry ADDS to inventory.available_qty.
CREATE TABLE IF NOT EXISTS receipt_log (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date           TEXT    NOT NULL,
    material_code        TEXT    NOT NULL,
    material_name        TEXT,
    uom                  TEXT,
    received_qty         REAL    NOT NULL DEFAULT 0,
    notes                TEXT,
    submitted_at         TEXT    DEFAULT (datetime('now')),
    FOREIGN KEY (material_code) REFERENCES inventory(material_code)
);
"""


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def load_clean_excel():
    """Load and clean all three Excel source files."""

    # ── File A: Inventory ────────────────────────────────────────────────────
    df_a = pd.read_excel(PATH_A, sheet_name=SHEET_A)
    df_a.columns = df_a.columns.str.strip()
    df_a["Material_Code"] = df_a["Material_Code"].astype(str).str.strip()
    df_a["Material_Name"] = df_a["Material_Name"].astype(str).str.strip()
    df_a["Available_Qty"] = pd.to_numeric(df_a["Available_Qty"], errors="coerce").fillna(0)

    # Detect ordered qty column
    ordered_col = None
    for c in df_a.columns:
        if c.strip() in ("Ordered_Qty", "Balance To Be Received"):
            ordered_col = c; break
    if ordered_col:
        df_a[ordered_col] = pd.to_numeric(df_a[ordered_col], errors="coerce").fillna(0)
    else:
        df_a["Ordered_Qty"] = 0.0
        ordered_col = "Ordered_Qty"

    inv = df_a.groupby("Material_Code", as_index=False).agg(
        Material_Name =("Material_Name", "first"),
        Nature        =("Nature",        "first"),
        UOM           =("UOM",           "first"),
        Available_Qty =("Available_Qty", "sum"),
        Ordered_Qty   =(ordered_col,     "sum"),
    )

    # ── File B: Recipe ───────────────────────────────────────────────────────
    df_b = pd.read_excel(PATH_B, sheet_name=SHEET_B)
    df_b.columns = df_b.columns.str.strip()
    df_b = df_b.dropna(subset=["Lining_System_Code", "Material_Code"])
    df_b["Material_Code"] = df_b["Material_Code"].astype(str).str.strip()
    df_b = df_b.assign(
        Material_Code=df_b["Material_Code"].str.split(r",\s*")
    ).explode("Material_Code").reset_index(drop=True)
    df_b["Material_Code"] = df_b["Material_Code"].str.strip()
    df_b["Lining_System_Code"] = (
        df_b["Lining_System_Code"].astype(float).astype(int).astype(str)
    )
    df_b["For_1_SQM"] = pd.to_numeric(df_b["For_1_SQM"], errors="coerce")
    recipe = df_b[[
        "Lining_System_Code", "Lining_System_Short_Name", "Lining_Type",
        "Material_Code", "Material_Description", "Material_Name", "For_1_SQM", "UOM"
    ]].copy()

    # ── File C: Equipment ────────────────────────────────────────────────────
    df_c = pd.read_excel(PATH_C, sheet_name=SHEET_C)
    df_c.columns = df_c.columns.str.strip()
    df_c = df_c.dropna(subset=["Equipment_Tag_No.", "Lining_System_Code"])
    df_c["Equipment_Tag_No."] = df_c["Equipment_Tag_No."].astype(str).str.strip()
    df_c["Location"]           = df_c["Location"].astype(str).str.strip()
    df_c["Type"]               = df_c["Type"].astype(str).str.strip()
    df_c["Surface_Area_SQM"]   = pd.to_numeric(df_c["Surface_Area_SQM"], errors="coerce")
    df_c["Lining_System_Code"] = (
        df_c["Lining_System_Code"].astype(float).astype(int).astype(str)
    )
    # Strip whitespace from all string columns
    for col in ["Name", "Description", "Lining_System_Short_Name"]:
        if col in df_c.columns:
            df_c[col] = df_c[col].astype(str).str.strip()
    equip = df_c[[
        "Location", "Type", "Lining_System_Code", "Lining_System_Short_Name",
        "Lining_Type", "Equipment_Tag_No.", "Name", "Description",
        "Material Spec.", "Design", "Surface_Area_SQM", "Lining_System+"
    ]].copy()

    return inv, recipe, equip


def seed_database(conn, inv, recipe, equip):
    """Insert master data. Safe to re-run: clears and reloads master tables,
    but preserves consumption_log and sqm_progress."""
    cur = conn.cursor()

    # ── Rebuild master tables ─────────────────────────────────────────────
    cur.execute("DELETE FROM inventory")
    cur.execute("DELETE FROM recipe")
    cur.execute("DELETE FROM equipment")

    # inventory
    for _, r in inv.iterrows():
        cur.execute("""
            INSERT INTO inventory (material_code, material_name, nature, uom,
                                   available_qty, ordered_qty)
            VALUES (?,?,?,?,?,?)
        """, (r["Material_Code"], r["Material_Name"], r.get("Nature",""),
              r["UOM"], r["Available_Qty"], r.get("Ordered_Qty", 0)))

    # recipe
    for _, r in recipe.iterrows():
        cur.execute("""
            INSERT INTO recipe (lining_system_code, lining_system_short_name,
                                lining_type, material_code, material_description,
                                material_name, for_1_sqm, uom)
            VALUES (?,?,?,?,?,?,?,?)
        """, (r["Lining_System_Code"], r.get("Lining_System_Short_Name",""),
              r.get("Lining_Type",""), r["Material_Code"],
              r.get("Material_Description",""), r.get("Material_Name",""),
              r.get("For_1_SQM", 0), r.get("UOM","")))

    # equipment
    for _, r in equip.iterrows():
        cur.execute("""
            INSERT INTO equipment (location, type, lining_system_code,
                                   lining_system_short_name, lining_type,
                                   equipment_tag, name, description,
                                   material_spec, design, surface_area_sqm,
                                   lining_systems)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (r["Location"], r["Type"], r["Lining_System_Code"],
              r.get("Lining_System_Short_Name",""), r.get("Lining_Type",""),
              r["Equipment_Tag_No."], r.get("Name",""), r.get("Description",""),
              r.get("Material Spec.",""), r.get("Design",""),
              r.get("Surface_Area_SQM", 0), r.get("Lining_System+","")))

    # ── Seed sqm_progress (insert only new rows, preserve existing done_sqm) ─
    equip_sc = equip.groupby(
        ["Equipment_Tag_No.", "Lining_System_Code"], as_index=False
    )["Surface_Area_SQM"].sum()

    for _, r in equip_sc.iterrows():
        cur.execute("""
            INSERT INTO sqm_progress (equipment_tag, lining_system_code,
                                      original_sqm, done_sqm)
            VALUES (?,?,?,0)
            ON CONFLICT(equipment_tag, lining_system_code)
            DO UPDATE SET original_sqm = excluded.original_sqm
        """, (r["Equipment_Tag_No."], r["Lining_System_Code"],
              r["Surface_Area_SQM"]))

    conn.commit()
    print(f"  ✅ inventory     : {len(inv)} rows")
    print(f"  ✅ recipe        : {len(recipe)} rows")
    print(f"  ✅ equipment     : {len(equip)} rows")
    print(f"  ✅ sqm_progress  : {len(equip_sc)} (tag, code) pairs seeded")


def main():
    print("=" * 60)
    print("  Smart Material Estimator — Database Setup")
    print("=" * 60)
    print(f"\n  Database: {DB_PATH}\n")

    print("  Loading and cleaning Excel files…")
    inv, recipe, equip = load_clean_excel()

    print("  Creating schema…")
    conn = connect()
    conn.executescript(SCHEMA)
    conn.commit()

    print("  Seeding master data…")
    seed_database(conn, inv, recipe, equip)

    # Summary
    cur = conn.cursor()
    for tbl in ["inventory","recipe","equipment","sqm_progress","consumption_log","receipt_log"]:
        n = cur.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
        print(f"    {tbl:<22}: {n} rows")

    conn.close()
    print("\n  ✅ sme_database.db is ready.\n")
    print("  Next step: run  streamlit run app.py")


if __name__ == "__main__":
    main()