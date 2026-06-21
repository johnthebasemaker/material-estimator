"""
Smart Material Estimator — Data Loading & Validation Script
============================================================
Purpose : Load the three source Excel files, clean known data-quality
          issues, and verify that all three DataFrames can be joined
          without silent data loss.

Run     : python validate_data.py
          (paths below are defaults; Streamlit app will receive these
           as BytesIO objects from st.file_uploader instead)
"""

import pandas as pd
import numpy as np
import sys

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — RAW LOAD
# ─────────────────────────────────────────────────────────────────────────────

FILE_A = "/mnt/project/Materials_DetailsAvailable_Qty.xlsx"
FILE_B = "/mnt/project/For_1_SQM.xlsx"
FILE_C = "/mnt/project/Equipment.xlsx"


def load_raw():
    df_a = pd.read_excel(FILE_A, sheet_name="Materials")
    df_b = pd.read_excel(FILE_B, sheet_name="LINING SYSTEM MATERIAL CONSM")
    df_c = pd.read_excel(FILE_C, sheet_name="Data Input")
    return df_a, df_b, df_c


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — CLEAN & NORMALISE
# ─────────────────────────────────────────────────────────────────────────────

def clean_inventory(df_a: pd.DataFrame) -> pd.DataFrame:
    """
    FIX 1 — Duplicate Material_Code rows (same material split across two POs).
    Resolution : sum Available_Qty per Material_Code; keep first occurrence of
                 descriptive columns (Material_Name, Nature, UOM).
    """
    df = df_a.copy()

    # Keep only the columns needed downstream
    # Keep Ordered_Qty if present — load_all will aggregate it
    keep = ["Material_Code", "Material_Name", "Nature", "UOM", "Available_Qty"]
    for col_candidate in ["Ordered_Qty", "Balance To Be Received"]:
        if col_candidate in df.columns:
            keep.append(col_candidate)
            break
    df = df[keep].copy()

    # Strip whitespace from string columns
    df["Material_Code"] = df["Material_Code"].astype(str).str.strip()
    df["Material_Name"] = df["Material_Name"].astype(str).str.strip()

    # Aggregate duplicates
    agg = df.groupby("Material_Code", as_index=False).agg(
        Material_Name=("Material_Name", "first"),
        Nature=("Nature", "first"),
        UOM=("UOM", "first"),
        Available_Qty=("Available_Qty", "sum"),
    )
    return agg


def clean_recipe(df_b: pd.DataFrame) -> pd.DataFrame:
    """
    FIX 2 — Material_Code in File B has two comma-separated multi-code cells
             ('GI-8005766, GI-8005765' and 'GI-8005766, GI-8005767').
    Resolution : explode these rows so each row holds exactly one Material_Code.

    FIX 3 — Lining_System_Code stored as float64 (1.0, 2.0 …).
    Resolution : cast to int then str for consistent join keys.
    """
    df = df_b.copy()

    keep = [
        "Lining_System_Code", "Lining_System_Short_Name", "Lining_Type",
        "Material_Code", "Material_Description", "Material_Name",
        "For_1_SQM", "UOM",
    ]
    df = df[keep].copy()

    # Drop rows where both keys are null (section-header spacer rows)
    df = df.dropna(subset=["Lining_System_Code", "Material_Code"])

    # Explode comma-separated Material_Code cells
    df["Material_Code"] = df["Material_Code"].astype(str).str.strip()
    df = df.assign(
        Material_Code=df["Material_Code"].str.split(r",\s*")
    ).explode("Material_Code").reset_index(drop=True)
    df["Material_Code"] = df["Material_Code"].str.strip()

    # Normalise Lining_System_Code → integer string  e.g. "1", "2"
    df["Lining_System_Code"] = (
        df["Lining_System_Code"].astype(float).astype(int).astype(str)
    )

    # Numeric guard for For_1_SQM
    df["For_1_SQM"] = pd.to_numeric(df["For_1_SQM"], errors="coerce")

    return df


def clean_equipment(df_c: pd.DataFrame) -> pd.DataFrame:
    """
    FIX 4 — Surface_Area_SQM is object dtype (mixed int / float strings).
    Resolution : coerce to numeric; rows with NaN area are non-lining rows
                 (structural entries) and are dropped.

    FIX 5 — Equipment_Tag_No. has 120 NaN rows (sub-header / blank rows) and
             mixed types (int 7112 vs str '513-37213-AGI-501').
    Resolution : drop rows with null Tag_No.; cast all to string.
    """
    df = df_c.copy()

    keep = [
        "Location", "Type", "Lining_System_Code", "Lining_System_Short_Name",
        "Lining_Type", "Equipment_Tag_No.", "Name", "Description",
        "Surface_Area_SQM",
    ]
    df = df[keep].copy()

    # Drop completely blank / header rows (must be before type coercions)
    df = df.dropna(subset=["Equipment_Tag_No.", "Lining_System_Code"])

    # Strip whitespace from all string columns  (Name has trailing spaces in source)
    for col in ["Equipment_Tag_No.", "Name", "Description", "Location",
                "Lining_System_Short_Name", "Lining_Type"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Normalise Lining_System_Code → integer string (same as File B)
    df["Lining_System_Code"] = (
        df["Lining_System_Code"].astype(float).astype(int).astype(str)
    )

    # Coerce Surface_Area_SQM to float; drop rows where it is null / zero
    df["Surface_Area_SQM"] = pd.to_numeric(df["Surface_Area_SQM"], errors="coerce")
    before = len(df)
    df = df.dropna(subset=["Surface_Area_SQM"])
    df = df[df["Surface_Area_SQM"] > 0]
    after = len(df)
    if before != after:
        print(
            f"  [WARN] Equipment: dropped {before - after} rows with "
            f"null/zero Surface_Area_SQM"
        )

    return df.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — VALIDATION SUITE
# ─────────────────────────────────────────────────────────────────────────────

PASS = "  ✅ PASS"
FAIL = "  ❌ FAIL"
WARN = "  ⚠️  WARN"


def validate(inv: pd.DataFrame, recipe: pd.DataFrame, equip: pd.DataFrame):
    errors = []
    warnings = []

    print("\n" + "=" * 65)
    print("  SMART MATERIAL ESTIMATOR — DATA VALIDATION REPORT")
    print("=" * 65)

    # ── Shape checks ────────────────────────────────────────────────
    print("\n[1] Shape after cleaning")
    print(f"  File A (Inventory)  : {inv.shape[0]:>4} rows × {inv.shape[1]} cols")
    print(f"  File B (Recipe)     : {recipe.shape[0]:>4} rows × {recipe.shape[1]} cols")
    print(f"  File C (Equipment)  : {equip.shape[0]:>4} rows × {equip.shape[1]} cols")

    # ── Null checks on join keys ─────────────────────────────────────
    print("\n[2] Null checks on primary join keys")
    for label, df, col in [
        ("File A • Material_Code",       inv,    "Material_Code"),
        ("File B • Material_Code",        recipe, "Material_Code"),
        ("File B • Lining_System_Code",   recipe, "Lining_System_Code"),
        ("File C • Lining_System_Code",   equip,  "Lining_System_Code"),
        ("File C • Equipment_Tag_No.",    equip,  "Equipment_Tag_No."),
        ("File C • Surface_Area_SQM",     equip,  "Surface_Area_SQM"),
    ]:
        n_null = df[col].isna().sum()
        status = PASS if n_null == 0 else FAIL
        if n_null > 0:
            errors.append(f"Nulls in {label}: {n_null}")
        print(f"{status}  {label:<40}  nulls={n_null}")

    # ── Duplicate Material_Code in inventory ─────────────────────────
    print("\n[3] Duplicate Material_Code in cleaned inventory")
    dups = inv["Material_Code"].duplicated().sum()
    status = PASS if dups == 0 else FAIL
    if dups > 0:
        errors.append(f"Still {dups} duplicate Material_Code rows after aggregation")
    print(f"{status}  Duplicate codes remaining : {dups}")

    # ── Lining_System_Code join coverage ────────────────────────────
    print("\n[4] Lining_System_Code join coverage (File C → File B)")
    lsc_c = set(equip["Lining_System_Code"].unique())
    lsc_b = set(recipe["Lining_System_Code"].unique())
    unmatched = lsc_c - lsc_b
    status = PASS if not unmatched else FAIL
    if unmatched:
        errors.append(f"Unmatched Lining_System_Codes in Equipment: {unmatched}")
    print(f"{status}  Equipment codes with no recipe : {unmatched or 'none'}")
    extra_b = lsc_b - lsc_c
    if extra_b:
        warnings.append(f"Recipe codes not used by any equipment: {extra_b}")
        print(f"{WARN}  Recipe codes unused by equipment : {extra_b}")

    # ── Material_Code join coverage ──────────────────────────────────
    print("\n[5] Material_Code join coverage (File B → File A)")
    mc_b = set(recipe["Material_Code"].unique())
    mc_a = set(inv["Material_Code"].unique())
    missing_inv = mc_b - mc_a
    status = PASS if not missing_inv else WARN
    if missing_inv:
        warnings.append(f"Recipe materials absent from inventory: {missing_inv}")
    print(f"{status}  Recipe materials with no inventory record : "
          f"{missing_inv or 'none'}")

    # ── Test merge: Equipment × Recipe ──────────────────────────────
    print("\n[6] Test merge — Equipment × Recipe (on Lining_System_Code)")
    merged_er = pd.merge(
        equip, recipe, on="Lining_System_Code", how="left",
        suffixes=("_equip", "_recipe")
    )
    null_material = merged_er["Material_Code"].isna().sum()
    status = PASS if null_material == 0 else FAIL
    if null_material > 0:
        errors.append(
            f"Equipment×Recipe merge produced {null_material} rows "
            f"with null Material_Code (missing recipe entries)"
        )
    print(f"{status}  Merged rows     : {len(merged_er)}")
    print(f"  Equipment input rows  : {len(equip)}")
    print(f"  Null Material_Code after merge : {null_material}")

    # ── Test merge: (Equipment × Recipe) × Inventory ────────────────
    print("\n[7] Test merge — Full 3-way join")
    merged_full = pd.merge(
        merged_er, inv, on="Material_Code", how="left",
        suffixes=("", "_inv")
    )
    null_avail = merged_full["Available_Qty"].isna().sum()
    status = PASS if null_avail == 0 else WARN
    if null_avail > 0:
        warnings.append(
            f"Full join has {null_avail} rows with no Available_Qty "
            f"(materials in recipe but not in inventory)"
        )
    print(f"{status}  Final joined rows     : {len(merged_full)}")
    print(f"  Rows with Available_Qty = NaN : {null_avail}")
    print(f"  Unique equipment tags covered : "
          f"{merged_full['Equipment_Tag_No.'].nunique()}")

    # ── Quick demand preview ─────────────────────────────────────────
    print("\n[8] Quick demand preview (top 5 materials by total demand)")
    merged_full["Demand"] = (
        merged_full["For_1_SQM"] * merged_full["Surface_Area_SQM"]
    )
    # After 3-way merge, Material_Name may come from recipe (_x) or inv (_y)
    mat_name_col = next(
        c for c in ["Material_Name", "Material_Name_x", "Material_Name_y"]
        if c in merged_full.columns
    )
    preview = (
        merged_full.groupby(["Material_Code", mat_name_col], as_index=False)
        ["Demand"].sum()
        .sort_values("Demand", ascending=False)
        .head(5)
    )
    preview.columns = ["Material_Code", "Material_Name", "Total_Demand"]
    print(preview.to_string(index=False))

    # ── Summary ──────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  SUMMARY")
    print("=" * 65)
    if errors:
        print(f"\n  ❌ {len(errors)} ERROR(s) found — must fix before app launch:")
        for e in errors:
            print(f"     • {e}")
    else:
        print("\n  ✅ Zero blocking errors detected.")

    if warnings:
        print(f"\n  ⚠️  {len(warnings)} WARNING(s) — review recommended:")
        for w in warnings:
            print(f"     • {w}")
    else:
        print("  ✅ Zero warnings.")

    print("\n  DataFrames are ready for the Streamlit application.\n")
    return len(errors) == 0, merged_full


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("\nLoading raw Excel files …")
    df_a_raw, df_b_raw, df_c_raw = load_raw()
    print("  Loaded File A :", df_a_raw.shape)
    print("  Loaded File B :", df_b_raw.shape)
    print("  Loaded File C :", df_c_raw.shape)

    print("\nCleaning & normalising …")
    inv     = clean_inventory(df_a_raw)
    recipe  = clean_recipe(df_b_raw)
    equip   = clean_equipment(df_c_raw)

    ok, full_df = validate(inv, recipe, equip)
    sys.exit(0 if ok else 1)