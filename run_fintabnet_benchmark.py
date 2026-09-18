# -*- coding: utf-8 -*-
"""Run our extractor against a random sample of real FinTabNet.c tables and
score the result with tablekit.scorer -- the actual external, published
benchmark (as opposed to run_ground_truth_suite.py's own small hand-picked
set). See fintabnet_synthesize.py's docstring for why this works from
synthesized single-page PDFs rather than table images: FinTabNet.c's public
annotation archive doesn't ship the original PDFs (nor does any actively
maintained mirror), but its ground truth already records every cell's real
PDF-coordinate position and text -- enough to faithfully reconstruct the
same text layer without needing the original file.

Usage:
    python run_fintabnet_benchmark.py [--n 50] [--seed 0] [--split val]
"""
import argparse
import json
import random
import sys
from pathlib import Path

import extract_all_tables as X
from fintabnet_synthesize import synthesize_pdf, build_ground_truth_rows
from tablekit.scorer import score_table

ROOT = Path(__file__).resolve().parent
ANNOT_DIR = ROOT / "fintabnet" / "FinTabNet.c-PDF_Annotations"
SCRATCH_DIR = ROOT / "fintabnet" / "_synth"


def sample_tables(split, n, rng):
    """Randomly sample up to `n` tables matching `split` ('train'/'val'/
    'test'/None for all), WITHOUT parsing every one of the ~77k annotation
    files up front -- shuffle the (cheap) file listing first, then parse
    files one at a time only until enough matching tables are found. A file
    can hold more than one table on the same page, and not every table in
    a file necessarily matches the requested split, so this over-samples
    files rather than assuming a 1:1 file:table ratio."""
    files = list(ANNOT_DIR.glob("*_tables.json"))
    rng.shuffle(files)
    out = []
    for jf in files:
        if len(out) >= n:
            break
        try:
            tables = json.loads(jf.read_text(encoding="utf-8"))
        except Exception:
            continue
        for i, t in enumerate(tables):
            if split is None or t.get("split") == split:
                out.append((jf, i, t))
    rng.shuffle(out)
    return out[:n]


def best_matching_table(tables, gt_rows):
    if not tables:
        return None
    return min(tables, key=lambda t: abs(len(t["rows"]) - len(gt_rows)))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--n", type=int, default=50, help="sample size")
    ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--split", default="val", choices=["train", "val", "test", "all"])
    args = ap.parse_args(argv)

    if not ANNOT_DIR.exists():
        print(f"Annotation dir not found: {ANNOT_DIR}")
        return 1

    split = None if args.split == "all" else args.split
    rng = random.Random(args.seed)
    sample = sample_tables(split, args.n, rng)
    print(f"sampled {len(sample)} tables (split={args.split!r})")

    SCRATCH_DIR.mkdir(parents=True, exist_ok=True)
    scores = []
    for jf, idx, t in sample:
        case_name = f"{jf.stem}_{idx}"
        gt_rows = build_ground_truth_rows(t)
        if not gt_rows or not any(any(c for c in r) for r in gt_rows):
            continue
        pdf_path = SCRATCH_DIR / f"{case_name}.pdf"
        try:
            synthesize_pdf(t, pdf_path)
        except Exception as e:
            scores.append((case_name, 0.0, f"synthesis failed: {e}"))
            continue
        try:
            tables = X.scan([pdf_path], {0}, min_rows=1, min_cols=1)
        except Exception as e:
            scores.append((case_name, 0.0, f"scan failed: {e}"))
            continue
        match = best_matching_table(tables, gt_rows)
        if match is None:
            scores.append((case_name, 0.0, "no table found"))
            continue
        result = score_table(gt_rows, match["rows"])
        scores.append((case_name, result["score"], None))

    scores.sort(key=lambda s: s[1])
    print(f"\n{'case':<45} {'score':>7}  detail")
    print("-" * 70)
    for name, score, note in scores:
        print(f"{name:<45} {score:>7.3f}  {note or ''}")

    if scores:
        avg = sum(s[1] for s in scores) / len(scores)
        misses = sum(1 for _, s, _ in scores if s == 0.0)
        perfect = sum(1 for _, s, _ in scores if s >= 0.999)
        print("-" * 70)
        print(f"AVERAGE over {len(scores)} tables: {avg:.3f}  "
              f"({perfect} perfect, {misses} complete misses)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
