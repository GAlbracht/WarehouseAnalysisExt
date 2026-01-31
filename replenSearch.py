

from __future__ import annotations

from pathlib import Path
import re
import pandas as pd


# ---------------- CONFIG ----------------
INPUT_FILE = r"physicalInventory.xlsx"   # change if needed
SHEET_NAME = 0                           # 0 = first sheet; or set "Sheet1"

OUT_DIR = Path("material_location_with_bins_v4")
OUT_DIR.mkdir(parents=True, exist_ok=True)

OUT_XLSX = OUT_DIR / "material_location_report.xlsx"
OUT_CSV_LOC = OUT_DIR / "material_locations.csv"
OUT_CSV_CONS = OUT_DIR / "consolidation_opportunities.csv"

# Expected Excel column names
COL_BIN = "Storage Bin"
COL_MAT = "Product"
COL_HU = "Handling Unit"
COL_BATCH = "Batch"
COL_QTY = "Quantity"  # optional
COL_GR_DATE = "Goods Receipt Date"
COL_GR_TIME = "Goods Receipt Time"
COL_STORAGE_TYPE = "Storage Type"
COL_STOCK_TYPE = "Stock Type"

VALID_STORAGE_TYPES = {"S015", "S012"}
VALID_STOCK_TYPE = "F2"


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


def classify_zone(bin_val: str) -> str:
    """A is filtered out; classify remaining."""
    if not bin_val:
        return "unknown"
    b = bin_val.upper()
    if b.startswith("B-"):
        return "b"
    if b.startswith("C-"):
        return "c"
    if "PLT" in b:
        return "plt"
    return "unknown"


def auto_fit_columns(ws, df: pd.DataFrame, min_w: int = 10, max_w: int = 60):
    """Approximate column widths."""
    def excel_col(n: int) -> str:
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    for i, col in enumerate(df.columns, start=1):
        header_len = len(str(col))
        sample = df[col].astype(str).head(200).tolist()
        max_len = max([header_len] + [len(v) for v in sample if v is not None])
        ws.column_dimensions[excel_col(i)].width = max(min_w, min(max_w, max_len + 2))


# ---------------- MAIN ----------------
def main():
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    df.columns = [clean_text(c) for c in df.columns]

    required = [
        COL_BIN, COL_MAT, COL_HU, COL_BATCH,
        COL_GR_DATE, COL_GR_TIME,
        COL_STORAGE_TYPE, COL_STOCK_TYPE,
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}\nFound: {list(df.columns)}")

    # Clean fields
    df[COL_BIN] = df[COL_BIN].apply(norm_bin)
    df[COL_MAT] = df[COL_MAT].apply(clean_text)
    df[COL_HU] = df[COL_HU].apply(clean_handling_unit)
    df[COL_BATCH] = df[COL_BATCH].apply(clean_text)
    df[COL_STORAGE_TYPE] = df[COL_STORAGE_TYPE].apply(clean_text).str.upper()
    df[COL_STOCK_TYPE] = df[COL_STOCK_TYPE].apply(clean_text).str.upper()

    # Filters
    df = df[df[COL_STORAGE_TYPE].isin(VALID_STORAGE_TYPES) & (df[COL_STOCK_TYPE] == VALID_STOCK_TYPE)]
    df = df[~df[COL_BIN].str.startswith("A-", na=False)]  # filter out A
    df = df[(df[COL_MAT] != "") & (df[COL_BIN] != "")]

    # Parse date/time for sorting
    df["_gr_date"] = pd.to_datetime(df[COL_GR_DATE], errors="coerce")
    df["_gr_time"] = pd.to_datetime(df[COL_GR_TIME], errors="coerce").dt.time

    # --- Count distinct HUs per (material, bin, batch) ---
    hu_counts = (
        df.groupby([COL_MAT, COL_BIN, COL_BATCH])[COL_HU]
        .nunique(dropna=True)
        .reset_index(name="hu_count_same_bin_batch")
    )

    # --- Pick ONE representative HU per (material, bin, batch) ---
    # "Best" = most recent GR Date/Time within that bin+batch group.
    df_sorted_for_rep = df.sort_values(
        by=[COL_MAT, COL_BIN, COL_BATCH, "_gr_date", "_gr_time"],
        ascending=[True, True, True, False, False],
    )
    rep = (
        df_sorted_for_rep.groupby([COL_MAT, COL_BIN, COL_BATCH], as_index=False)
        .first()
    )

    # Join HU count + zone
    rep = rep.merge(hu_counts, on=[COL_MAT, COL_BIN, COL_BATCH], how="left")
    rep["zone"] = rep[COL_BIN].apply(classify_zone)
    rep["can_consolidate"] = rep["hu_count_same_bin_batch"].apply(lambda n: "YES" if (pd.notna(n) and n > 1) else "")

    # Keep one row per bin+batch (rep HU)
    # Order options per material: GR date desc, GR time desc, batch asc
    rep = rep.sort_values(
        by=["_gr_date", "_gr_time", COL_BATCH],
        ascending=[True, True, True]
    ).copy()

    rep["fifo_rank"] = rep.groupby(COL_MAT).cumcount() + 1

    # Output columns
    cols_out = [
        COL_MAT,
        "zone",
        "fifo_rank",
        COL_BIN,
        COL_BATCH,
        COL_HU,  # representative HU
        "hu_count_same_bin_batch",
        COL_GR_DATE,
        COL_GR_TIME,
    ]
    if COL_QTY in rep.columns:
        cols_out.append(COL_QTY)
    cols_out += [COL_STORAGE_TYPE, COL_STOCK_TYPE, "can_consolidate"]

    material_locations = rep[cols_out].copy()

    # Consolidation sheet = only rows needing consolidation
    consolidate = material_locations[material_locations["hu_count_same_bin_batch"] > 1].copy()
    consolidate["note"] = "CONSOLIDATE (multiple HUs same bin+batch)"

    # Nice ordering: zone then material then rank
    zone_rank = {"b": 1, "c": 2, "plt": 3, "unknown": 4}
    material_locations["_zone_rank"] = material_locations["zone"].map(zone_rank).fillna(99).astype(int)
    material_locations = (
        material_locations.sort_values(by=["_zone_rank", COL_MAT, "fifo_rank"], ascending=[True, True, True])
                         .drop(columns=["_zone_rank"])
    )

    consolidate["_zone_rank"] = consolidate["zone"].map(zone_rank).fillna(99).astype(int)
    consolidate = (
        consolidate.sort_values(by=["_zone_rank", COL_MAT, "fifo_rank"], ascending=[True, True, True])
                   .drop(columns=["_zone_rank"])
    )

    # --- CSVs for quick checks ---
    material_locations.to_csv(OUT_CSV_LOC, index=False)
    consolidate.to_csv(OUT_CSV_CONS, index=False)

    # --- One workbook, two sheets ---
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        material_locations.to_excel(writer, sheet_name="Material_Locations", index=False)
        consolidate.to_excel(writer, sheet_name="Consolidation", index=False)

        wb = writer.book

        ws1 = wb["Material_Locations"]
        ws1.freeze_panes = "A2"
        ws1.auto_filter.ref = ws1.dimensions
        for cell in ws1[1]:
            cell.font = cell.font.copy(bold=True)

        ws2 = wb["Consolidation"]
        ws2.freeze_panes = "A2"
        ws2.auto_filter.ref = ws2.dimensions
        for cell in ws2[1]:
            cell.font = cell.font.copy(bold=True)

        try:
            auto_fit_columns(ws1, material_locations)
            auto_fit_columns(ws2, consolidate)
        except Exception:
            pass

    print("Done.")
    print(f"- Excel workbook: {OUT_XLSX}")
    print(f"- CSV (material locations): {OUT_CSV_LOC}")
    print(f"- CSV (consolidation opportunities): {OUT_CSV_CONS}")
    print("\nWhat changed:")
    print("- Each material/bin/batch now shows ONE representative HU plus HU count for that bin+batch.")
    print("- Consolidation sheet shows where HU_count_same_bin_batch > 1.")


if __name__ == "__main__":
    main()
