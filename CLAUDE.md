# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

Converts NinjaRMM device-list CSV exports (49 columns) into a focused 17-column XLSX report with cleaned/normalized data. Used by MSP staff to prepare client-facing device inventory reports.

## Running the Script

```bash
pip install pandas openpyxl

python3 clean-ninja-export.py <input.csv> [output.xlsx]
# Default output: <input>_cleaned.xlsx in the same directory
```

## Architecture

Single-file pipeline in `clean-ninja-export.py`:

1. **`KEEP_COLUMNS`** (line 21) — defines the 17 output columns and their order
2. **`process()`** (line 235) — core pipeline: read CSV → validate columns → clean → write XLSX
3. **`_write_xlsx()`** (line 260) — Excel styling (blue headers, frozen pane, autofilter, auto-sized columns)

### Cleaning functions (each handles NaN safely):
- `clean_processor()` — strips Intel/AMD boilerplate, deduplicates multi-socket CPUs (`2× Intel Xeon Silver 4309Y`)
- `clean_volumes()` — extracts C: drive, formats as `C: 75.3 / 235.5 GiB free (68% used)`
- `clean_os_name()` — collapses verbose OS strings (`Microsoft Windows 11 Pro 10.0.22631` → `Win 11 Pro`)
- `clean_memory()` — rounds reported RAM to nearest standard tier (8, 16, 32, 64 GiB, etc.)

### Input tolerance
- Warns on missing columns but continues processing
- Ignores extra input columns beyond the 17 kept
