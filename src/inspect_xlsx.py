"""
Dev utility: dump the structure of a workbook so we can map exact cells.
Usage:  .venv/bin/python src/inspect_xlsx.py <file.xlsx> [max_rows] [max_cols]
Prints, per sheet: dimensions and the top-left block of values AND formulas.
"""
import sys
from pathlib import Path

import openpyxl


def dump(path: str, max_rows: int = 25, max_cols: int = 12):
    p = Path(path)
    print(f"\n=== {p.name} ===")
    vals = openpyxl.load_workbook(p, data_only=True)
    fmls = openpyxl.load_workbook(p, data_only=False)
    print("Sheets:", vals.sheetnames)
    for name in vals.sheetnames:
        ws_v, ws_f = vals[name], fmls[name]
        print(f"\n--- sheet: {name!r}  dims={ws_v.dimensions}  "
              f"max_row={ws_v.max_row} max_col={ws_v.max_column} ---")
        for r in range(1, min(ws_v.max_row, max_rows) + 1):
            cells = []
            for c in range(1, min(ws_v.max_column, max_cols) + 1):
                v = ws_v.cell(r, c).value
                f = ws_f.cell(r, c).value
                coord = ws_v.cell(r, c).coordinate
                if v is None and f is None:
                    continue
                if isinstance(f, str) and f.startswith("="):
                    cells.append(f"{coord}=[formula]{f} -> {v!r}")
                else:
                    cells.append(f"{coord}={v!r}")
            if cells:
                print(f"  r{r}: " + " | ".join(cells))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    args = sys.argv[1:]
    mr = int(args[1]) if len(args) > 1 else 25
    mc = int(args[2]) if len(args) > 2 else 12
    dump(args[0], mr, mc)
