from __future__ import annotations

from pathlib import Path
from typing import Optional, List, Tuple, Union
import pandas as pd


PICK_FILE = r"Pick Area Organization(The One To Rule Them All).xlsx"
PHYSICAL_FILE = r"physicalInventory.xlsx"

# Set these to a sheet name if you want (otherwise first sheet is used)
PICK_SHEET: Optional[Union[str, int]] = None
PHYSICAL_SHEET: Optional[Union[str, int]] = None

OUT_DIR = Path("full_mismatch_results")


def clean_str(x) -> Optional[str]:
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None


def clean_bin(x) -> Optional[str]:
    s = clean_str(x)
    return s.upper() if s else None


def read_first_sheet_if_none(path: str | Path, sheet: Optional[Union[str, int]]) -> pd.DataFrame:
    """
    If sheet is None, read the first worksheet.
    If sheet is provided (name or index), read that sheet.
    """
    if sheet is None:
        # Read first sheet explicitly by index 0
        return pd.read_excel(path, sheet_name=0)
    return pd.read_excel(path, sheet_name=sheet)


def compare_bins(
    pick_path: str | Path,
    physical_path: str | Path,
    pick_sheet: Optional[Union[str, int]] = None,
    physical_sheet: Optional[Union[str, int]] = None,
    out_dir: Path = OUT_DIR,
) -> Tuple[pd.DataFrame, List[str]]:

    out_dir.mkdir(parents=True, exist_ok=True)

    # Load (guaranteed DataFrames)
    pick_df = read_first_sheet_if_none(pick_path, pick_sheet)
    phys_df = read_first_sheet_if_none(physical_path, physical_sheet)

    # Normalize headers
    pick_df.columns = [str(c).strip().lower() for c in pick_df.columns]
    phys_df.columns = [str(c).strip().lower() for c in phys_df.columns]

    # REQUIRED columns (exact names you gave)
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
    pick.columns = ["storage_bin", "pick_material"]
    pick["storage_bin"] = pick["storage_bin"].apply(clean_bin)
    pick["pick_material"] = pick["pick_material"].apply(clean_str)
    pick = pick.dropna(subset=["storage_bin"])

    phys = phys_df[[PHYS_BIN_COL, PHYS_PROD_COL]].copy()
    phys.columns = ["storage_bin", "physical_product"]
    phys["storage_bin"] = phys["storage_bin"].apply(clean_bin)
    phys["physical_product"] = phys["physical_product"].apply(clean_str)
    phys = phys.dropna(subset=["storage_bin"])

    # Aggregate physical products per bin (handles multiple rows per bin)
    phys_agg = (
        phys.dropna(subset=["physical_product"])
        .groupby("storage_bin")["physical_product"]
        .apply(lambda s: sorted(set(s.tolist())))
        .reset_index(name="physical_products_list")
    )

    # Merge: keep all pick bins
    merged = pick.merge(phys_agg, on="storage_bin", how="left")

    def mismatch_type(row) -> Optional[str]:
        pick_mat = row["pick_material"]
        phys_list = row["physical_products_list"]

        # No physical record for that bin
        if not isinstance(phys_list, list):
            return "Missing in Physical Inventory"

        # If pick material blank, ignore (change to mismatch if you want)
        if pick_mat is None:
            return None

        return None if pick_mat in phys_list else "Product/Material mismatch"

    merged["mismatch_type"] = merged.apply(mismatch_type, axis=1)

    mismatches = merged[merged["mismatch_type"].notna()].copy()
    mismatches["physical_products"] = mismatches["physical_products_list"].apply(
        lambda v: ", ".join(v) if isinstance(v, list) else ""
    )
    mismatches.drop(columns=["physical_products_list"], inplace=True)

    mismatched_bins = sorted(mismatches["storage_bin"].dropna().unique().tolist())

    # Save outputs
    (out_dir / "mismatch_bins.txt").write_text("\n".join(mismatched_bins), encoding="utf-8")
    mismatches.to_csv(out_dir / "mismatches.csv", index=False)
    mismatches.to_excel(out_dir / "mismatches.xlsx", index=False)

    return mismatches, mismatched_bins


if __name__ == "__main__":
    mismatches_df, bins = compare_bins(
        pick_path=PICK_FILE,
        physical_path=PHYSICAL_FILE,
        pick_sheet=PICK_SHEET,
        physical_sheet=PHYSICAL_SHEET,
        out_dir=OUT_DIR,
    )

    print(f"Total mismatched bins: {len(bins)}")
    print("First 25 mismatched bins:")
    for b in bins[:25]:
        print(b)

    print("\nSaved:")
    print(f" - {OUT_DIR / 'mismatch_bins.txt'}")
    print(f" - {OUT_DIR / 'mismatches.xlsx'}")
    print(f" - {OUT_DIR / 'mismatches.csv'}")
