# -*- coding: utf-8 -*-
"""
Created on Mon Nov 24 13:44:05 2025

@author: AnatoliiPurchel
"""

import os
import re
import sys
import argparse
from pathlib import Path
import pandas as pd
import numpy as np

Number = r"-?\d+(?:\.\d+)?"

def parse_one_file(path: Path):
    """
    Parse a Waters ASCII file.

    RULE:
      - FUNCTION 1 => treat numeric pairs as (mz, intensity)   [MS]
      - FUNCTION != 1 => treat numeric pairs as (channel, intensity) [PDA/UV]
        even though these blocks also include "Scan" lines.
    """
    rows = []
    fn = None              # current FUNCTION (int)
    rt = None              # current retention time (float)
    scan = None            # current scan number (int)

    re_function = re.compile(r"^\s*FUNCTION\s+(\d+)\s*$", re.I)
    re_scan     = re.compile(r"^\s*Scan\s+(\d+)\s*$", re.I)
    re_rt       = re.compile(r"^\s*Retention\s+Time\s+(" + Number + r")\s*$", re.I)
    re_pair     = re.compile(r"^\s*(" + Number + r")\s+(" + Number + r")\s*$")

    with path.open("r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            m = re_function.match(line)
            if m:
                fn = int(m.group(1))
                rt = None
                scan = None
                continue

            m = re_scan.match(line)
            if m:
                scan = int(m.group(1))
                continue

            m = re_rt.match(line)
            if m:
                rt = float(m.group(1))
                continue

            m = re_pair.match(line)
            if m and fn is not None and rt is not None:
                a = float(m.group(1))
                b = float(m.group(2))

                if fn == 1:
                    # MS data: (mz, intensity)
                    rows.append({
                        "function": fn,
                        "scan": scan,
                        "time": rt,
                        "mz": a,
                        "intensity": b,
                        "channel_id": "MS",
                        "channel_label": "MS",
                        "source_file": path.name,
                        "source_path": str(path),
                    })
                else:
                    # PDA / UV channel data: (channel_number, intensity)
                    ch_str = f"{a:.4f}"
                    rows.append({
                        "function": fn,
                        "scan": scan,
                        "time": rt,
                        "mz": None,
                        "intensity": b,
                        "channel_id": ch_str,    # e.g., "220.0000"
                        "channel_label": ch_str,
                        "source_file": path.name,
                        "source_path": str(path),
                    })
    return rows


def iter_txt_files(root: Path, recursive: bool):
    """
   Yield .txt files under root that are expected to be Waters MassLynx ASCII exports.
   """
    return root.rglob("*.txt") if recursive else root.glob("*.txt")


def safe_sheet_name(name: str, used: set):
    # Excel rules: max 31 chars; no : \ / ? * [ ]
    bad = ':/\\?*[]'
    for ch in bad:
        name = name.replace(ch, "_")
    base = name[:31]
    cand = base
    i = 1
    while cand in used:
        suf = f"_{i}"
        cand = base[:31-len(suf)] + suf
        i += 1
    used.add(cand)
    return cand


def build_wide_sheet(df: pd.DataFrame, time_decimals: int = 3) -> pd.DataFrame:
    """
    For one (function, channel_id) subset:
      - Build union time grid (rounded)
      - For each source_file, create two columns: <file>_time, <file>_intensity
    """
    df = df.copy()
    df["time_round"] = df["time"].round(time_decimals)

    grid = np.sort(df["time_round"].unique())
    wide = pd.DataFrame(index=grid)

    for src, sub in df.groupby("source_file"):
        sub = sub.sort_values(["time_round", "time"]).drop_duplicates("time_round", keep="first")
        sub = sub.set_index("time_round")

        time_col = pd.Series(sub["time"], index=sub.index).reindex(grid)
        inten_col = pd.Series(sub["intensity"], index=sub.index).reindex(grid)

        wide[f"{src}_time"] = time_col.values
        wide[f"{src}_intensity"] = inten_col.values

    return wide.reset_index(drop=True)


def main():
    ap = argparse.ArgumentParser(
        description="Parse Waters ASCII files -> one Excel, one worksheet per channel, with 2 columns per chromatogram (time/intensity)."
    )
    ap.add_argument("input_dir", help="Folder containing .txt files")
    ap.add_argument("-o", "--output", default="combined_by_channel_wide.xlsx",
                    help="Output Excel filename (default: combined_by_channel_wide.xlsx)")
    ap.add_argument("--recursive", action="store_true", help="Scan subfolders recursively")
    ap.add_argument("--sheet-prefix", default="", help="Optional prefix for sheet names (e.g., 'Day1_')")
    ap.add_argument("--time-decimals", type=int, default=3,
                    help="Decimals for rounding Retention Time to align grids (default: 3)")
    args = ap.parse_args()

    root = Path(args.input_dir).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        print(f"[ERROR] Not a directory: {root}")
        sys.exit(1)

    files = list(iter_txt_files(root, args.recursive))
    if not files:
        print(f"[INFO] No .txt files found under {root} (recursive={args.recursive}).")
        sys.exit(0)

    print(f"[INFO] Found {len(files)} file(s). Parsing...")
    all_rows = []
    for fp in files:
        rows = parse_one_file(fp)
        if rows:
            all_rows.extend(rows)
            print(f"  ✓ {fp.name}: {len(rows)} row(s)")
        else:
            print(f"  ⚠ {fp.name}: no recognizable traces found")

    if not all_rows:
        print("[ERROR] No data parsed from any files. Nothing to write.")
        sys.exit(1)

    df = pd.DataFrame(all_rows)

    out_path = (root / args.output).resolve()
    used_names = set()

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        # Build per-channel sheets
        index_rows = []
        for (func, chan_id), sub in df.groupby(["function", "channel_id"], sort=True):
            # Human-friendly label:
            if func == 1:
                label = "MS"
            else:
                # show channel number without trailing .0000 when possible
                try:
                    ch_num = float(chan_id)
                    ch_disp = str(int(round(ch_num))) if abs(ch_num - round(ch_num)) < 1e-6 else f"{ch_num:g}"
                except Exception:
                    ch_disp = str(chan_id)
                label = f"ch-{ch_disp}"

            sheet_base = f"{args.sheet_prefix}F{func}_{label}"
            sheet_name = safe_sheet_name(sheet_base, used_names)

            wide = build_wide_sheet(sub[["source_file", "time", "intensity"]].copy(),
                                    time_decimals=args.time_decimals)
            wide.to_excel(xw, sheet_name=sheet_name, index=False)

            index_rows.append({
                "sheet_name": sheet_name,
                "function": func,
                "channel_id": chan_id,
                "channel_label": label,
                "chromatograms": sub["source_file"].nunique(),
                "rows_in_sheet": len(wide)
            })

        # INDEX sheet
        idx_df = pd.DataFrame(index_rows).sort_values(["function", "channel_label", "sheet_name"])
        idx_df.to_excel(xw, sheet_name="INDEX", index=False)

    print(f"[SUMMARY] Wrote workbook: {out_path}")


if __name__ == "__main__":
    # Tell the script which folder to scan by default
    # Example: sys.argv += [r"C:\MassLynx Data\MassLynx\TXT files"]
    # Add this if you want it to search subfolders too:
    # sys.argv += ["--recursive"]
    main()
