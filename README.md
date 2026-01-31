# Warehouse Audit & A-Zone Replenishment Tools

This repository contains two Python scripts designed to support warehouse operations by:
- Auditing pick bin assignments against physical inventory
- Generating a FIFO-based A-zone replenishment plan for missing or mismatched materials

Both scripts are built for Excel/CSV inputs and produce analyst-friendly outputs.

---

## Scripts Overview

### `fullaudit.py` — Pick vs Physical Inventory Audit

**Purpose**  
Compares the Pick Area Organization against Physical Inventory to identify bin-level mismatches.

**Mismatch conditions**
A mismatch is flagged if:
- A storage bin exists in the pick layout but is missing in physical inventory, or
- A material assigned to a pick bin is not found among physical products for that bin

**Inputs**
| File | Required Columns |
|----|----|
| Pick Area Organization | `storage bin`, `material` |
| Physical Inventory | `storage bin`, `product` |

Column names are matched case-insensitively.

**Outputs** (`./output_mismatch_results/`)
- `mismatch_bins.txt` – bins with mismatches
- `mismatches.xlsx` – detailed mismatch report
- `mismatches.csv` – CSV version of the mismatch report

**Use cases**
- Pick area audits
- Bin accuracy KPIs
- Consolidation and cleanup analysis
- Pre-replenishment validation

---

### `replen_from_A_with_fifo.py` — A-Zone FIFO Replenishment Planner

**Purpose**  
Identifies materials expected in A-zone pick bins that are missing or mismatched in physical inventory, then recommends the top 5 FIFO reserve picks for replenishment.

This script converts audit findings into actionable replenishment decisions.

---

## Business Logic

### A-Zone Definition
- A-bins are storage bins starting with `A-` (case-insensitive)

### Replenishment Criteria
A material requires replenishment if:
- It is assigned to an A-bin in the Pick Area Organization
- It is missing or mismatched in Physical Inventory

### FIFO Source Filtering
Eligible reserve inventory must meet all of the following:
- `Storage Type` in `{S015, S012}`
- `Stock Type` = `F2`
- `Storage Bin` does not start with `A-`

### FIFO Ranking
- FIFO order is calculated using Goods Receipt Date and Time
- Higher `fifo_rank` indicates older inventory
- Top 5 FIFO candidates are selected per material

### Handling Unit Logic
- One representative handling unit is selected per `(bin, batch)`
- Handling unit count for the same `(bin, batch)` is included for context

---

## Inputs

| File | Required Columns |
|----|----|
| Pick Area Organization (The One To Rule Them All).xlsx | `storage bin`, `material` |
| physicalInventory.xlsx | `storage bin`, `product`, `handling unit`, `batch`, `goods receipt date`, `goods receipt time`, `storage type`, `stock type` |

Column names are case-insensitive but must exist.

---

## Outputs

### Folder: `./a_replenishment_plan/`

**Excel**
- `a_replenishment_plan.xlsx`
  - `A_Needs`: A-bin lines requiring replenishment
  - `Top5_FIFO_Picks`: Top 5 FIFO reserve options per material

**CSV**
- `top5_fifo_picks.csv`

---

## Installation

```bash
pip install pandas openpyxl
