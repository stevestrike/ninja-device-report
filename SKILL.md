# Ninja Device Report Cleanup

Converts a raw NinjaRMM device-list CSV export into a client-friendly XLSX report by keeping only the 17 relevant columns and cleaning up the noisy ones (CPU strings, volume strings, OS names, memory rounding).

## What the script does

**Keeps 17 columns, in this order** (original header names preserved):
Organization, Location, Display Name, Type, Device Role, Policy, Last Update, Warranty Start Date_formatted, Warranty End Date_formatted, Last Login, Memory Capacity GiB, OS Name, System Name, Device Model, Serial Number, Processors Name, Volumes

**Cleans four columns:**

- **Processors Name**: strips `(R)`, `(TM)`, `CPU`, and `@ N.NNGHz` clock specs. For multi-socket machines where the same CPU appears N times comma-separated, collapses to `N× <cpu name>`. Single-CPU machines get no count prefix. Handles AMD `with Radeon Graphics` trailers.
  - `Intel(R) Core(TM) i5-14500T` → `Intel Core i5-14500T`
  - `Intel(R) Xeon(R) CPU E5-2630 v3 @ 2.40GHz` → `Intel Xeon E5-2630 v3`
  - `Intel(R) Xeon(R) Silver 4309Y CPU @ 2.80GHz,Intel(R) Xeon(R) Silver 4309Y CPU @ 2.80GHz` → `2× Intel Xeon Silver 4309Y`

- **Volumes**: extracts only the `C:` drive, reformats as `C: <free> / <capacity> GiB free (<N>% used)`. Any non-C drives are dropped from the output. If no C: drive is present (rare), the cell is empty.
  - `Name: "C:"/ Type: "Local Disk"/ Capacity: "252850466816 (235.5 GiB)"/ Free: "82778816512 (77.1 GiB)"/ Usage %: "67%"` → `C: 77.1 / 235.5 GiB free (67% used)`

- **OS Name**: drops leading `Microsoft `, collapses `Windows` → `Win`, strips trailing build numbers.
  - `Microsoft Windows 11 Pro 10.0.22631` → `Win 11 Pro`

- **Memory Capacity GiB**: snaps reported RAM to the nearest standard purchase tier (1, 2, 4, 6, 8, 12, 16, 24, 32, 48, 64, 96, 128, 192, 256, 384, 512, 768, 1024, 1536, 2048, 3072, 4096 GB). Ninja reports *usable* memory, which sits slightly below installed capacity due to BIOS reserve and integrated graphics — e.g., a 32GB laptop may report 31.42 GiB. The script rounds up to the tier if the reading is within 5% below it, otherwise falls back to nearest-tier.
  - `15.87` → `16`
  - `31.42` → `32`
  - `255.5` → `256`
  - `511.5` → `512`

**Output formatting:**
- XLSX with a header row (white text on dark-blue fill, bold)
- Frozen top row
- AutoFilter enabled
- Column widths auto-sized to content (capped at 45 chars)
- Calibri 11pt headers, Calibri 10pt body

## Updating the script

If NinjaRMM changes their export format (new columns, renamed columns, different CPU strings), edit `scripts/clean_ninja_export.py`:

- Column list: `KEEP_COLUMNS` at the top of the file
- CPU rules: `clean_processor` and `_clean_single_cpu`
- Volumes rules: `clean_volumes` and `_format_volume`
- OS rules: `clean_os_name`
- Memory tiers: `STANDARD_MEMORY_TIERS` in `clean_memory`

The script tolerates missing columns — it prints a warning to stderr and skips them rather than failing. New unexpected columns are simply ignored (they're not in `KEEP_COLUMNS`).
