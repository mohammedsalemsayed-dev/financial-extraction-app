# -*- coding: utf-8 -*-
"""Classify each FinTabNet ground-truth table as a real financial-figure
table (numbers dominate the non-label cells) vs. prose/legal content (exhibit
indices, narrative text tables) that this tool was never built to parse --
so a raw average across ALL of FinTabNet doesn't distinguish "the tool is
bad at real financial tables" from "the sample includes tables outside the
tool's stated purpose." Reports both splits' scores separately."""
import json
import random
import sys
from pathlib import Path

import extract_all_tables as X
from fintabnet_synthesize import synthesize_pdf, build_ground_truth_rows
from tablekit.parse import parse_number
from tablekit.scorer import score_table
from run_fintabnet_benchmark import ANNOT_DIR, SCRATCH_DIR, sample_tables, best_matching_table


def is_financial(gt_rows, threshold=0.35):
    cells = [c for r in gt_rows for c in r if c]
    if not cells:
        return False
    numeric = sum(1 for c in cells if parse_number(c) is not None)
    return numeric / len(cells) >= threshold


def main():
    n = int(sys.argv[sys.argv.index("--n") + 1]) if "--n" in sys.argv else 200
    seed = int(sys.argv[sys.argv.index("--seed") + 1]) if "--seed" in sys.argv else 2
    rng = random.Random(seed)
    sample = sample_tables("val", n, rng)
    SCRATCH_DIR.mkdir(parents=True, exist_ok=True)

    fin_scores, other_scores = [], []
    for jf, idx, t in sample:
        case_name = f"{jf.stem}_{idx}"
        gt_rows = build_ground_truth_rows(t)
        if not gt_rows or not any(any(c for c in r) for r in gt_rows):
            continue
        pdf_path = SCRATCH_DIR / f"{case_name}.pdf"
        try:
            synthesize_pdf(t, pdf_path)
            tables = X.scan([pdf_path], {0}, min_rows=1, min_cols=1)
        except Exception:
            continue
        match = best_matching_table(tables, gt_rows)
        score = score_table(gt_rows, match["rows"])["score"] if match else 0.0
        (fin_scores if is_financial(gt_rows) else other_scores).append(score)

    for label, scores in [("financial-figure tables", fin_scores),
                           ("prose/legal/other tables", other_scores)]:
        if scores:
            avg = sum(scores) / len(scores)
            print(f"{label:<28} n={len(scores):<4} avg={avg:.3f}")
        else:
            print(f"{label:<28} n=0")
    total = len(fin_scores) + len(other_scores)
    print(f"\n{len(fin_scores)}/{total} ({100*len(fin_scores)/total:.0f}%) of sampled tables are financial-figure tables")


if __name__ == "__main__":
    main()
