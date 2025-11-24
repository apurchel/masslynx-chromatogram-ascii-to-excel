# masslynx-chromatogram-ascii-to-excel
Parse Waters MassLynx data exported into ASCII .txt files into wide-format Excel for chromatogram plotting (per function/channel).
Python script to parse **Waters MassLynx ASCII `.txt` exports** and combine them into a single **Excel workbook**:

- One worksheet per **(function, channel)** (MS or PDA/UV)
- Each worksheet in **wide format** with:
  - Two columns per chromatogram: `<source_file>_time`, `<source_file>_intensity`
  - A shared time grid (rounded retention time)

This is useful for downstream plotting and analysis of chromatograms in Excel, Python, R, etc.

---

## Features

- Supports Waters **FUNCTION 1** (MS) and **FUNCTION != 1** (PDA/UV) data.
- Automatically discovers `.txt` files in a directory (optionally recursive).
- Creates one Excel worksheet per channel, plus an **INDEX** worksheet.
- Handles non-uniform time grids by building a **union time grid** per channel and aligning traces.
- Avoids sheet name collisions and Excel’s sheet name length/character limits.

---

## Input format

This script expects **Waters MassLynx ASCII export** text files.

Parsing rules:

- Lines like `FUNCTION 1` define the current function.
- Lines like `Scan 123` define the scan number (optional metadata).
- Lines like `Retention Time 3.456` define the retention time (in minutes).
- Numeric pairs then follow:

  - For `FUNCTION 1`: numeric pairs are interpreted as **(mz, intensity)**.
  - For `FUNCTION != 1`: numeric pairs are interpreted as **(channel_number, intensity)** (PDA/UV).

Only numeric pairs that appear **after** a FUNCTION and Retention Time are used.

---

## Output

The script writes a single Excel file (default `combined_by_channel_wide.xlsx`) into the input directory.

- One worksheet per **(function, channel_id)**:
  - Sheet names look like `F1_MS`, `F2_ch-220`, etc.
  - You can optionally prefix sheet names (e.g. `Day1_F1_MS`).

Each per-channel sheet contains:

- Columns: for each source file, two columns:
  - `<source_file>_time`
  - `<source_file>_intensity`
- Rows: union of rounded retention times across all files for that channel.

There is also an **INDEX** sheet summarizing:

- `sheet_name`
- `function`
- `channel_id`
- `channel_label`
- `chromatograms` (number of input files contributing)
- `rows_in_sheet`

---

## Installation

Clone the repo:

```bash
git clone https://github.com/<your-username>/masslynx-ascii-chromatograms.git
cd masslynx-ascii-chromatograms
