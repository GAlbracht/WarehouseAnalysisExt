# Warehouse Inventory Audit & Replenishment Analytics

This repository contains Python-based analytics tools designed to support warehouse inventory accuracy, replenishment planning, and reserve stock optimization.

The scripts operate on real operational inventory data and implement business rules commonly used in warehouse and supply chain environments, including FIFO logic, bin validation, and consolidation analysis.

---

## Repository Contents

### `fullaudit.py`
Audits Pick Area Organization data against Physical Inventory to identify bin- and material-level mismatches.

Used to:
- Validate pick bin assignments
- Identify missing or mismatched inventory
- Support bin accuracy KPIs and cleanup initiatives

---

### `replen_from_A_with_fifo.py`
An end-to-end A-zone replenishment planner that combines multiple audit and replenishment workflows into a single script.

Used to:
- Identify materials expected in A-zone pick bins that are missing or mismatched
- Filter eligible reserve inventory
- Rank FIFO replenishment options using goods receipt date and time
- Output the top 5 FIFO picks per material

This script consolidates earlier, separate audit and FIFO-ranking logic into one operational workflow.

---

### `replenSearch.py`
Material location and consolidation classifier for reserve inventory.

Used to:
- Locate reserve inventory by material, bin, and batch
- Rank available locations for picking
- Identify consolidation opportunities where multiple handling units exist in the same bin and batch

---

## Technical Highlights

- Python (pandas, openpyxl)
- Inventory mismatch detection and validation logic
- FIFO ranking using timestamp-based aging
- Data aggregation at bin, batch, and material levels
- Excel-first reporting for operational teams

---

## Outputs

All scripts generate Excel and CSV outputs designed for:
- Operations review
- Replenishment execution
- KPI tracking
- Further analysis in Excel or BI tools

---

## Installation

```bash
pip install pandas openpyxl
