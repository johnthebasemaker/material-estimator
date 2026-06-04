"""
Smart Material Estimator — Dynamic setup_db.py
===============================================
Syncs Excel data to SQLite. Automatically detects new columns
added to Excel and adds them to the SQLite database without losing data.
"""

import os, sqlite3
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "sme_database.db")
PATH_A   = os.path.join(BASE_DIR, "Materials_DetailsAvailable_Qty.xlsx")
PATH_B   = os.path.join(BASE_DIR, "For_1_SQM.xlsx")
PATH_C   = os.path.join(BASE_DIR, "Equipment.xlsx")

SCHEMA = """
CREATE TABLE IF NOT EXISTS inventory (material_code TEXT PRIMARY KEY);
CREATE TABLE IF NOT EXISTS recipe (id INTEGER PRIMARY KEY AUTOINCREMENT, lining_system_code TEXT NOT NULL, material_code TEXT NOT NULL, FOREIGN KEY (material_code) REFERENCES inventory(material_code));
CREATE TABLE IF NOT EXISTS equipment (id INTEGER PRIMARY KEY AUTOINCREMENT, equipment_tag TEXT NOT NULL, lining_system_code TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS sqm_progress (equipment_tag TEXT NOT NULL, lining_system_code TEXT NOT NULL, original_sqm REAL NOT NULL DEFAULT 0, done_sqm REAL NOT NULL DEFAULT 0, PRIMARY KEY (equipment_tag, lining_system_code));
CREATE TABLE IF NOT EXISTS consumption_log (id INTEGER PRIMARY KEY AUTOINCREMENT, entry_date TEXT NOT NULL, equipment_tag TEXT NOT NULL, lining_system_code TEXT NOT NULL, sqm_completed REAL NOT NULL DEFAULT 0, material_code TEXT NOT NULL, expected_qty REAL, consumed_qty REAL NOT NULL DEFAULT 0, submitted_at TEXT DEFAULT (datetime('now')));
CREATE TABLE IF NOT EXISTS receipt_log (id INTEGER PRIMARY KEY AUTOINCREMENT, entry_date TEXT NOT NULL, material_code TEXT NOT NULL, received_qty REAL NOT NULL DEFAULT 0, submitted_at TEXT DEFAULT (datetime('now')), FOREIGN KEY (material_code) REFERENCES inventory(material_code));
CREATE TABLE IF NOT EXISTS draft_consumption (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    session_key         TEXT    NOT NULL,
    entry_date          TEXT    NOT NULL,
    equipment_tag       TEXT    NOT NULL,
    lining_system_code  TEXT    NOT NULL,
    lining_system_name  TEXT,
    sqm_completed       REAL    NOT NULL DEFAULT 0,
    material_code       TEXT    NOT NULL,
    material_name       TEXT,
    uom                 TEXT,
    expected_qty        REAL,
    actual_qty          REAL    NOT NULL DEFAULT 0,
    effective_qty       REAL,
    variance_pct        REAL,
    variance_status     TEXT,
    notes               TEXT,
    added_at            TEXT    DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS orders_log (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id            TEXT    NOT NULL,
    order_date          TEXT    NOT NULL,
    generated_by        TEXT    DEFAULT 'Smart Material Estimator',
    generation_source   TEXT,
    source_detail       TEXT,
    material_code       TEXT    NOT NULL,
    material_name       TEXT,
    uom                 TEXT,
    ordered_qty         REAL    NOT NULL DEFAULT 0,
    fulfilled_qty       REAL    NOT NULL DEFAULT 0,
    status              TEXT    DEFAULT 'Pending',
    receipt_ids         TEXT,
    notes               TEXT,
    submitted_at        TEXT    DEFAULT (datetime('now'))
);
"""


def _migrate_existing_db(conn):
    """Apply non-destructive migrations to an existing database."""
    cur = conn.cursor()
    # Add order_id column to receipt_log if missing
    existing = {r[1].lower() for r in cur.execute("PRAGMA table_info(receipt_log)").fetchall()}
    if "order_id" not in existing:
        cur.execute("ALTER TABLE receipt_log ADD COLUMN order_id TEXT")
        print("  [+] Migrated receipt_log: added order_id column")
    # Add lining_system_name and variance columns to consumption_log if missing
    cl_cols = {r[1].lower() for r in cur.execute("PRAGMA table_info(consumption_log)").fetchall()}
    for col, dtype in [("lining_system_name", "TEXT"), ("variance_status", "TEXT"), ("variance_pct", "REAL")]:
        if col not in cl_cols:
            cur.execute(f"ALTER TABLE consumption_log ADD COLUMN {col} {dtype}")
            print(f"  [+] Migrated consumption_log: added {col} column")
    ol_cols = {r[1].lower() for r in cur.execute("PRAGMA table_info(orders_log)").fetchall()}
    if "pr_number" not in ol_cols:
        cur.execute("ALTER TABLE orders_log ADD COLUMN pr_number TEXT")
        print("  [+] Migrated orders_log: added pr_number column")
    conn.commit()

def dynamic_sync_table(conn, df, table_name, primary_keys):
    """Dynamically adds missing columns to the DB, then upserts data."""
    # 1. Clean up accidental Excel ghost columns (like 'Unnamed: 10')
    df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False, na=False)].copy()
    
    # 2. Convert any Date/Time columns to strings so SQLite doesn't crash
    for col in df.select_dtypes(include=['datetime', 'datetimetz']).columns:
        df[col] = df[col].astype(str)
        
    cur = conn.cursor()
    
    # 3. Get existing DB columns and convert to lowercase for safe comparison
    existing_cols = {row[1].lower() for row in cur.execute(f"PRAGMA table_info({table_name})").fetchall()}
    pk_lower = [pk.lower() for pk in primary_keys]
    
    # 4. Add missing columns from Excel DataFrame to DB
    for col in df.columns:
        if col.lower() not in existing_cols and col.lower() not in pk_lower:
            dtype = "REAL" if pd.api.types.is_numeric_dtype(df[col]) else "TEXT"
            # FIX: Use double quotes around the column name to safely handle apostrophes
            cur.execute(f'ALTER TABLE {table_name} ADD COLUMN "{col}" {dtype}')
            print(f"  [+] Added new column '{col}' to {table_name}")
            
    # 5. Clear old master data and insert new (Does not clear logs/progress)
    cur.execute(f"DELETE FROM {table_name}")
    
    # 6. Insert data dynamically
    # FIX: Use double quotes around the column names here as well
    cols = ",".join([f'"{c}"' for c in df.columns])
    placeholders = ",".join(["?"] * len(df.columns))
    sql = f"INSERT INTO {table_name} ({cols}) VALUES ({placeholders})"
    
    # Convert df to records, replacing NaN with None for SQLite
    records = df.where(pd.notnull(df), None).values.tolist()
    cur.executemany(sql, records)
    print(f"  ✅ {table_name:<15}: Synced {len(df)} rows.")

def load_and_clean_excel():
    # File A
    df_a = pd.read_excel(PATH_A, sheet_name="Materials")
    df_a.columns = df_a.columns.str.strip()
    df_a.rename(columns={"Material_Code": "material_code"}, inplace=True)
    df_a["material_code"] = df_a["material_code"].astype(str).str.strip()
    
    # ── FIX 1: Convert empty quantities to 0 so SQLite doesn't crash ──
    if "Available_Qty" in df_a.columns:
        df_a["Available_Qty"] = pd.to_numeric(df_a["Available_Qty"], errors="coerce").fillna(0.0)
    if "Ordered_Qty" in df_a.columns:
        df_a["Ordered_Qty"] = pd.to_numeric(df_a["Ordered_Qty"], errors="coerce").fillna(0.0)
        
    # ── FIX 2: Squash duplicate Material Codes to prevent UNIQUE constraint crash ──
    # Sums numeric columns (like Qty) and takes the first value for text/date columns
    agg_dict = {}
    for col in df_a.columns:
        if col == "material_code":
            continue
        elif pd.api.types.is_numeric_dtype(df_a[col]):
            agg_dict[col] = "sum"
        else:
            agg_dict[col] = "first"
            
    df_a = df_a.groupby("material_code", as_index=False).agg(agg_dict)
    # ────────────────────────────────────────────────────────────────
    
    # File B
    df_b = pd.read_excel(PATH_B, sheet_name="LINING SYSTEM MATERIAL CONSM")
    df_b.columns = df_b.columns.str.strip()
    df_b = df_b.dropna(subset=["Lining_System_Code", "Material_Code"])
    df_b.rename(columns={"Lining_System_Code": "lining_system_code", "Material_Code": "material_code"}, inplace=True)
    df_b["lining_system_code"] = df_b["lining_system_code"].astype(float).astype(int).astype(str)
    df_b = df_b.assign(material_code=df_b["material_code"].astype(str).str.split(r",\s*")).explode("material_code")
    df_b["material_code"] = df_b["material_code"].str.strip()

    # File C
    df_c = pd.read_excel(PATH_C, sheet_name="Data Input")
    df_c.columns = df_c.columns.str.strip()
    df_c = df_c.dropna(subset=["Equipment_Tag_No.", "Lining_System_Code"])
    df_c.rename(columns={"Equipment_Tag_No.": "equipment_tag", "Lining_System_Code": "lining_system_code"}, inplace=True)
    df_c["equipment_tag"] = df_c["equipment_tag"].astype(str).str.strip()
    df_c["lining_system_code"] = df_c["lining_system_code"].astype(float).astype(int).astype(str)
    
    return df_a, df_b, df_c

def main():
    print("=" * 60 + "\n  Smart Material Estimator — Dynamic Database Setup\n" + "=" * 60)
    df_a, df_b, df_c = load_and_clean_excel()
    
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.executescript(SCHEMA)
    _migrate_existing_db(conn)
    
    dynamic_sync_table(conn, df_a, "inventory", ["material_code"])
    dynamic_sync_table(conn, df_b, "recipe", ["id"])
    dynamic_sync_table(conn, df_c, "equipment", ["id"])
    
    # Seed SQM Progress (Preserves existing done_sqm)
    sqm_data = df_c.groupby(["equipment_tag", "lining_system_code"], as_index=False)["Surface_Area_SQM"].sum()
    for _, r in sqm_data.iterrows():
        conn.execute("""
            INSERT INTO sqm_progress (equipment_tag, lining_system_code, original_sqm, done_sqm)
            VALUES (?,?,COALESCE(?, 0),0) ON CONFLICT(equipment_tag, lining_system_code)
            DO UPDATE SET original_sqm = excluded.original_sqm
        """, (r["equipment_tag"], r["lining_system_code"], r.get("Surface_Area_SQM", 0)))
    
    conn.commit()
    conn.close()
    print("\n  ✅ Database is ready and columns are synced. Run streamlit run app.py")

if __name__ == "__main__":
    main()