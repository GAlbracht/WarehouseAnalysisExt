
"""
A-zone Replenishment Planner (Combined)
--------------------------------------
Goal:
1) Identify materials that are expected in A bins (Pick Area Organization) but are missing/mismatched in Physical Inventory.
2) For each such material, provide the TOP 5 FIFO picks (oldest stock) to grab from reserve (non-A bins),
   ranked by fifo_rank where HIGHER = OLDER (as requested).

Inputs:
- Pick Area Organization(The One To Rule Them All).xlsx
  Required columns (case-insensitive): "storage bin", "material"
- physicalInventory.xlsx
  Required columns (case-insensitive): "storage bin", "product", plus columns for FIFO picking:
    "Storage Bin", "Product", "Handling Unit", "Batch",
    "Goods Receipt Date", "Goods Receipt Time", "Storage Type", "Stock Type"

Notes:
- "A bins" are identified as Storage Bin starting with "A-" (case-insensitive).
- FIFO candidate sources are filtered to:
    Storage Type in {"S015","S012"}
    Stock Type == "F2"
    Storage Bin NOT starting with "A-"
- FIFO ranking: within each material, fifo_rank is highest for the oldest GR Date/Time.
- For each (Material, Bin, Batch), we output a single representative HU plus the HU count in that same (bin, batch).

Outputs (./a_replenishment_plan):
- a_replenishment_plan.xlsx
    Sheet "A_Needs"          : A-bin lines that need replenishment (missing or mismatch)
    Sheet "Top5_FIFO_Picks"  : For each material, top 5 FIFO (oldest) source options to grab
- top5_fifo_picks.csv

Install:
    pip install pandas openpyxl
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Union
import re
import pandas as pd


# ---------------- CONFIG ----------------
PICK_FILE = r"Pick Area Organization(The One To Rule Them All).xlsx"
PHYSICAL_FILE = r"physicalInventory.xlsx"

# If None, read first sheet. Otherwise provide sheet name or index.
PICK_SHEET: Optional[Union[str, int]] = None
PHYSICAL_SHEET: Optional[Union[str, int]] = None

OUT_DIR = Path("a_replenishment_plan")
OUT_DIR.mkdir(parents=True, exist_ok=True)

OUT_XLSX = OUT_DIR / "a_replenishment_plan.xlsx"
OUT_NEEDS_CSV = OUT_DIR / "a_needs.csv"
OUT_PICKS_CSV = OUT_DIR / "top5_fifo_picks.csv"

# A-zone identification
A_PREFIX = "A-"

# FIFO filtering rules (same as your replenSearch V4)
VALID_STORAGE_TYPES = {"S015", "S012"}
VALID_STOCK_TYPE = "F2"

# Physical Inventory columns (FIFO source + compare)
# We normalize headers, but these are the expected display names in the physical workbook.
COL_BIN_STD = "Storage Bin"
COL_MAT_STD = "Product"
COL_HU_STD = "Handling Unit"
COL_BATCH_STD = "Batch"
COL_GR_DATE_STD = "Goods Receipt Date"
COL_GR_TIME_STD = "Goods Receipt Time"
COL_STORAGE_TYPE_STD = "Storage Type"
COL_STOCK_TYPE_STD = "Stock Type"


# ---------------- HELPERS ----------------
def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).replace("\u00A0", " ").strip()
    return " ".join(s.split())


def norm_bin(x) -> str:
    return clean_text(x).upper()


def clean_handling_unit(x) -> str:
    """Convert 3001328410.0 -> '3001328410'."""
    if pd.isna(x):
        return ""
    if isinstance(x, int):
        return str(x)
    if isinstance(x, float):
        if x.is_integer():
            return str(int(x))
        return format(x, "f").rstrip("0").rstrip(".")
    s = clean_text(x)
    if re.fullmatch(r"\d+\.0", s):
        return s[:-2]
    return s


def is_a_bin(bin_val: str) -> bool:
    b = (bin_val or "").upper()
    return b.startswith(A_PREFIX)


def classify_zone(bin_val: str) -> str:
    """Classify for nicer grouping in outputs."""
    b = (bin_val or "").upper()
    if b.startswith("A-"):
        return "a"
    if b.startswith("B-"):
        return "b"
    if b.startswith("C-"):
        return "c"
    if "PLT" in b:
        return "plt"
    return "unknown"


def read_first_sheet_if_none(path: str | Path, sheet: Optional[Union[str, int]]) -> pd.DataFrame:
    if sheet is None:
        return pd.read_excel(path, sheet_name=0)
    return pd.read_excel(path, sheet_name=sheet)


# ---------------- STEP 1: FIND A NEEDS ----------------
def find_a_replen_needs(
    pick_path: str | Path,
    physical_path: str | Path,
    pick_sheet: Optional[Union[str, int]] = None,
    physical_sheet: Optional[Union[str, int]] = None,
) -> pd.DataFrame:
    """
    Return rows for A bins where expected pick material is missing/mismatched in physical inventory.
    Output columns:
      a_bin, expected_material, mismatch_type, physical_products
    """
    pick_df = read_first_sheet_if_none(pick_path, pick_sheet)
    phys_df = read_first_sheet_if_none(physical_path, physical_sheet)

    # Normalize headers (lower for matching)
    pick_df.columns = [str(c).strip().lower() for c in pick_df.columns]
    phys_df.columns = [str(c).strip().lower() for c in phys_df.columns]

    # Required columns (as in fullAudit.py)
    PICK_BIN_COL = "storage bin"
    PICK_MAT_COL = "material"
    PHYS_BIN_COL = "storage bin"
    PHYS_PROD_COL = "product"

    missing = []
    for col in (PICK_BIN_COL, PICK_MAT_COL):
        if col not in pick_df.columns:
            missing.append(f"Pick file missing column: '{col}' (found: {list(pick_df.columns)})")
    for col in (PHYS_BIN_COL, PHYS_PROD_COL):
        if col not in phys_df.columns:
            missing.append(f"Physical file missing column: '{col}' (found: {list(phys_df.columns)})")
    if missing:
        raise ValueError("\n".join(missing))

    # Select + clean
    pick = pick_df[[PICK_BIN_COL, PICK_MAT_COL]].copy()
    pick.columns = ["storage_bin", "expected_material"]
    pick["storage_bin"] = pick["storage_bin"].apply(norm_bin)
    pick["expected_material"] = pick["expected_material"].apply(clean_text)
    pick = pick[(pick["storage_bin"] != "") & (pick["expected_material"] != "")]
    pick = pick[pick["storage_bin"].apply(is_a_bin)]  # only A bins

    phys = phys_df[[PHYS_BIN_COL, PHYS_PROD_COL]].copy()
    phys.columns = ["storage_bin", "physical_product"]
    phys["storage_bin"] = phys["storage_bin"].apply(norm_bin)
    phys["physical_product"] = phys["physical_product"].apply(clean_text)
    phys = phys[phys["storage_bin"] != ""]

    # Aggregate physical products per bin
    phys_agg = (
        phys[phys["physical_product"] != ""]
        .groupby("storage_bin")["physical_product"]
        .apply(lambda s: sorted(set(s.tolist())))
        .reset_index(name="physical_products_list")
    )

    merged = pick.merge(phys_agg, on="storage_bin", how="left")

    def mismatch_type(row) -> Optional[str]:
        exp_mat = row["expected_material"]
        phys_list = row["physical_products_list"]
        if not isinstance(phys_list, list):
            return "Missing in Physical Inventory"
        return None if exp_mat in phys_list else "Product/Material mismatch"

    merged["mismatch_type"] = merged.apply(mismatch_type, axis=1)
    needs = merged[merged["mismatch_type"].notna()].copy()
    needs["physical_products"] = needs["physical_products_list"].apply(
        lambda v: ", ".join(v) if isinstance(v, list) else ""
    )
    needs = needs.drop(columns=["physical_products_list"]).rename(columns={"storage_bin": "a_bin"})
    needs["zone"] = "a"
    needs = needs[["zone", "a_bin", "expected_material", "mismatch_type", "physical_products"]]

    # Stable sort
    needs = needs.sort_values(by=["a_bin", "expected_material"]).reset_index(drop=True)
    return needs


# ---------------- STEP 2: BUILD FIFO SOURCE OPTIONS ----------------
def build_fifo_source_options(
    physical_path: str | Path,
    physical_sheet: Optional[Union[str, int]] = None,
) -> pd.DataFrame:
    """
    Return one row per (material, bin, batch) for eligible source bins (non-A),
    with representative HU and fifo_rank where HIGHER = OLDER.
    """
    df = read_first_sheet_if_none(physical_path, physical_sheet)

    # Normalize columns by stripping but keep original names for expected set
    df.columns = [clean_text(c) for c in df.columns]

    required = [
        COL_BIN_STD, COL_MAT_STD, COL_HU_STD, COL_BATCH_STD,
        COL_GR_DATE_STD, COL_GR_TIME_STD, COL_STORAGE_TYPE_STD, COL_STOCK_TYPE_STD,
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns in physicalInventory for FIFO picking: {missing}\nFound: {list(df.columns)}")

    # Clean fields
    df[COL_BIN_STD] = df[COL_BIN_STD].apply(norm_bin)
    df[COL_MAT_STD] = df[COL_MAT_STD].apply(clean_text)
    df[COL_HU_STD] = df[COL_HU_STD].apply(clean_handling_unit)
    df[COL_BATCH_STD] = df[COL_BATCH_STD].apply(clean_text)
    df[COL_STORAGE_TYPE_STD] = df[COL_STORAGE_TYPE_STD].apply(clean_text).str.upper()
    df[COL_STOCK_TYPE_STD] = df[COL_STOCK_TYPE_STD].apply(clean_text).str.upper()

    # Filters for eligible source stock
    df = df[df[COL_STORAGE_TYPE_STD].isin(VALID_STORAGE_TYPES) & (df[COL_STOCK_TYPE_STD] == VALID_STOCK_TYPE)]
    df = df[~df[COL_BIN_STD].str.startswith("A-", na=False)]  # exclude A bins as sources
    df = df[(df[COL_MAT_STD] != "") & (df[COL_BIN_STD] != "")]

    # Parse date/time for sorting
    df["_gr_date"] = pd.to_datetime(df[COL_GR_DATE_STD], errors="coerce")
    df["_gr_time"] = pd.to_datetime(df[COL_GR_TIME_STD], errors="coerce").dt.time

    # Count distinct HUs per (material, bin, batch)
    hu_counts = (
        df.groupby([COL_MAT_STD, COL_BIN_STD, COL_BATCH_STD])[COL_HU_STD]
        .nunique(dropna=True)
        .reset_index(name="hu_count_same_bin_batch")
    )

    # Pick one representative HU per (material, bin, batch) by most recent GR date/time
    df_sorted_for_rep = df.sort_values(
        by=[COL_MAT_STD, COL_BIN_STD, COL_BATCH_STD, "_gr_date", "_gr_time"],
        ascending=[True, True, True, False, False],
    )
    rep = df_sorted_for_rep.groupby([COL_MAT_STD, COL_BIN_STD, COL_BATCH_STD], as_index=False).first()

    # Join HU count + zone
    rep = rep.merge(hu_counts, on=[COL_MAT_STD, COL_BIN_STD, COL_BATCH_STD], how="left")
    rep["zone"] = rep[COL_BIN_STD].apply(classify_zone)
    rep["can_consolidate"] = rep["hu_count_same_bin_batch"].apply(lambda n: "YES" if (pd.notna(n) and n > 1) else "")

    # Compute fifo_rank where HIGHER = OLDER
    # First, order newest -> oldest within each material
    rep = rep.sort_values(by=[COL_MAT_STD, "_gr_date", "_gr_time", COL_BATCH_STD], ascending=[True, False, False, True]).copy()
    rep["_rank_newest"] = rep.groupby(COL_MAT_STD).cumcount() + 1
    rep["_group_size"] = rep.groupby(COL_MAT_STD)["_rank_newest"].transform("max")
    rep["fifo_rank"] = rep["_group_size"] - rep["_rank_newest"] + 1  # oldest gets highest

    # Output columns
    cols_out = [
        COL_MAT_STD,
        "zone",
        "fifo_rank",
        COL_BIN_STD,
        COL_BATCH_STD,
        COL_HU_STD,  # representative HU
        "hu_count_same_bin_batch",
        COL_GR_DATE_STD,
        COL_GR_TIME_STD,
        COL_STORAGE_TYPE_STD,
        COL_STOCK_TYPE_STD,
        "can_consolidate",
    ]
    rep_out = rep[cols_out].copy()

    # Nice ordering: zone then material then fifo_rank desc (older first)
    zone_rank = {"b": 1, "c": 2, "plt": 3, "unknown": 4}
    rep_out["_zone_rank"] = rep_out["zone"].map(zone_rank).fillna(99).astype(int)
    rep_out = rep_out.sort_values(by=["_zone_rank", COL_MAT_STD, "fifo_rank"], ascending=[True, True, False]).drop(columns=["_zone_rank"])

    return rep_out.reset_index(drop=True)


# ---------------- STEP 3: TOP 5 PICKS PER MATERIAL ----------------
def top_fifo_picks_for_needs(needs_df: pd.DataFrame, fifo_sources_df: pd.DataFrame, top_n: int = 5) -> pd.DataFrame:
    """
    For each expected_material in needs_df, return top_n rows from fifo_sources_df by fifo_rank DESC (oldest first).
    Adds an 'pick_rank_within_material' (1..top_n) for readability.
    """
    mat_col_needs = "expected_material"
    mat_col_sources = COL_MAT_STD

    mats = sorted(set(needs_df[mat_col_needs].dropna().astype(str).tolist()))
    if not mats:
        return pd.DataFrame(columns=["expected_material", "pick_rank_within_material"] + fifo_sources_df.columns.tolist())

    sources_subset = fifo_sources_df[fifo_sources_df[mat_col_sources].isin(mats)].copy()
    if sources_subset.empty:
        # Still return empty with consistent columns
        out_cols = [mat_col_needs, "pick_rank_within_material"] + fifo_sources_df.columns.tolist()
        return pd.DataFrame(columns=out_cols)

    # For each material: pick oldest first => fifo_rank desc
    sources_subset = sources_subset.sort_values(by=[mat_col_sources, "fifo_rank"], ascending=[True, False])
    top = sources_subset.groupby(mat_col_sources, as_index=False, group_keys=False).head(top_n).copy()
    top["pick_rank_within_material"] = top.groupby(mat_col_sources).cumcount() + 1
    top = top.rename(columns={mat_col_sources: mat_col_needs})

    # Reorder columns
    cols = [
        mat_col_needs,
        "pick_rank_within_material",
        "zone",
        "fifo_rank",
        COL_BIN_STD,
        COL_BATCH_STD,
        COL_HU_STD,
        "hu_count_same_bin_batch",
        COL_GR_DATE_STD,
        COL_GR_TIME_STD,
        COL_STORAGE_TYPE_STD,
        COL_STOCK_TYPE_STD,
        "can_consolidate",
    ]
    # keep only existing cols (in case physical schema changes)
    cols = [c for c in cols if c in top.columns]
    return top[cols].reset_index(drop=True)


# ---------------- MAIN ----------------
def main():
    needs = find_a_replen_needs(
        pick_path=PICK_FILE,
        physical_path=PHYSICAL_FILE,
        pick_sheet=PICK_SHEET,
        physical_sheet=PHYSICAL_SHEET,
    )

    fifo_sources = build_fifo_source_options(
        physical_path=PHYSICAL_FILE,
        physical_sheet=PHYSICAL_SHEET,
    )

    top5 = top_fifo_picks_for_needs(needs, fifo_sources, top_n=5)

    # Save quick CSVs
    needs.to_csv(OUT_NEEDS_CSV, index=False)
    top5.to_csv(OUT_PICKS_CSV, index=False)

    # Save workbook
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        needs.to_excel(writer, sheet_name="A_Needs", index=False)
        top5.to_excel(writer, sheet_name="Top5_FIFO_Picks", index=False)

        wb = writer.book
        for sheet in ["A_Needs", "Top5_FIFO_Picks"]:
            ws = wb[sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)

    print("Done.")
    print(f"- Excel: {OUT_XLSX}")
    print(f"- CSV needs: {OUT_NEEDS_CSV}")
    print(f"- CSV picks: {OUT_PICKS_CSV}")

    if needs.empty:
        print("\nNo A-bin replenishment needs found (based on current compare rules).")
    else:
        print(f"\nA-bin replen needs rows: {len(needs)}")
        print(f"Unique materials to replenish: {needs['expected_material'].nunique()}")
        print("\nExample materials:", ", ".join(needs['expected_material'].dropna().unique().astype(str)[:10]))


if __name__ == "__main__":
    main()
