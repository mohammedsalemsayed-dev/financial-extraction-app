# -*- coding: utf-8 -*-
"""Score one extracted table against hand-transcribed ground truth.

Usage:
    python score_extraction.py ground_truth.json extracted.json

Each JSON file is just {"rows": [[...], [...], ...]} -- the same row-list
shape every other part of this project already uses. Prints an overall
score (0-1) and flags any row that didn't score a perfect match, so a
reviewer can jump straight to what's actually wrong instead of re-reading
the whole table.

This is a content-accuracy scorer (tablekit.scorer), not a TEDS/FinTabNet-
style structural benchmark -- see tablekit/scorer.py's docstring for why.
"""
import argparse
import json
import sys

from tablekit.scorer import score_table


def _load_rows(path):
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return data["rows"] if isinstance(data, dict) else data


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("ground_truth", help="hand-transcribed JSON: {\"rows\": [...]}")
    ap.add_argument("extracted", help="this tool's extracted JSON: {\"rows\": [...]}")
    ap.add_argument("--full", action="store_true",
                     help="print every row's detail, not just mismatches")
    args = ap.parse_args(argv)

    gt_rows = _load_rows(args.ground_truth)
    ex_rows = _load_rows(args.extracted)
    result = score_table(gt_rows, ex_rows)

    print(f"Overall score: {result['score']:.3f}  ({len(result['rows'])} rows aligned)")
    print()
    shown_any = False
    for i, row in enumerate(result["rows"]):
        if not args.full and row["row_score"] >= 0.999:
            continue
        shown_any = True
        label = (row.get("gt") or row.get("extracted") or [""])[0]
        note = row.get("note", "")
        print(f"  row {i}: score={row['row_score']:.3f}  {label!r}  {note}")
        if "cells" in row:
            for gt_cell, ex_cell in row["cells"]:
                if gt_cell != ex_cell:
                    print(f"      ground truth={gt_cell!r}  extracted={ex_cell!r}")
    if not shown_any and not args.full:
        print("  (every row matched perfectly)")
    return 0 if result["score"] >= 0.999 else 1


if __name__ == "__main__":
    sys.exit(main())
