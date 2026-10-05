#!/usr/bin/env python3
"""Show the structure of manually exported Tableau files (CSV/TSV/XLSX) so columns can be mapped.

Usage: python inspect_exports.py file1 [file2 ...] [--org "Org name"]

Prints encoding/delimiter guess, row count, columns, a few sample rows and the distinct values of
low-cardinality columns. With --org, rows are filtered to those where ANY cell equals the org name
(case-insensitive) and the filtered row count is shown. Nothing is modified or invented.
"""
import argparse
import csv
import io
import sys
from pathlib import Path


def read_rows(path):
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            sys.exit("openpyxl is required for .xlsx files: pip install openpyxl")
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        out = {}
        for ws in wb.worksheets:
            rows = [["" if c is None else str(c) for c in r] for r in ws.iter_rows(values_only=True)]
            out[f"sheet:{ws.title}"] = rows
        return out
    raw = p.read_bytes()
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            txt = raw.decode(enc)
            break
        except UnicodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(txt[:5000], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return {f"{enc} / delimiter={dialect.delimiter!r}": list(csv.reader(io.StringIO(txt), dialect))}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--org")
    a = ap.parse_args()
    for f in a.files:
        for label, rows in read_rows(f).items():
            print(f"\n=== {f} [{label}] ===")
            if not rows:
                print("(empty)")
                continue
            head, body = rows[0], rows[1:]
            print(f"{len(body)} data rows, {len(head)} columns")
            if a.org:
                body = [r for r in body if any(c.strip().lower() == a.org.strip().lower() for c in r)]
                print(f"rows matching org {a.org!r}: {len(body)}")
            for i, h in enumerate(head):
                vals = [r[i] for r in body if i < len(r) and r[i] != ""]
                uniq = sorted(set(vals))
                extra = f"distinct={len(uniq)}" + (f": {uniq[:12]}" if len(uniq) <= 12 else f" e.g. {uniq[:4]}")
                print(f"  [{i}] {h!r}  non-empty={len(vals)}  {extra}")
            print("  sample rows:")
            for r in body[:5]:
                print("   ", r)


if __name__ == "__main__":
    main()
