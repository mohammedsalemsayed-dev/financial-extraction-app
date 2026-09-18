# -*- coding: utf-8 -*-
"""Run every case in ground_truth/ through the real extractor and score it
against its hand-transcribed ground truth. This is the actual "get a real
number instead of a guess" step -- prints a summary table plus per-case
mismatch detail for anything that didn't score a perfect match.

Usage:
    python run_ground_truth_suite.py
"""
import json
from pathlib import Path

import extract_all_tables as X
from tablekit.parse import normspace
from tablekit.scorer import score_table

ROOT = Path(__file__).resolve().parent
GT_DIR = ROOT / "ground_truth"


def best_matching_table(tables, gt_rows, gt_title=None):
    """Among every table X.scan() found on the target page, pick the one
    that's actually the case's own table -- a stand-in for "the reviewer
    would have clicked this one," since scan() returns every table
    candidate on the page, not just the one we want.

    Prefer a title match when the ground truth carries one (every case in
    ground_truth/ does): falling back to "closest row count" alone breaks
    as soon as a SECOND, unrelated candidate happens to land nearer the
    target's row count than the real one -- found live once this project's
    extractor got better at finding note-level tables it used to miss
    entirely (eand2025_finance_costs.json: a newly-surfaced, genuinely
    correct 13-row tax-reconciliation table sat closer to this case's
    9-row ground truth than the real 4-row "Finance and other costs"
    table, so the old row-count-only picker silently started scoring the
    wrong candidate -- the real target's own score never moved)."""
    if not tables:
        return None
    if gt_title:
        needle = gt_title.strip().lower()
        title_hits = [t for t in tables if needle in (t.get("title") or "").strip().lower()]
        if len(title_hits) == 1:
            return title_hits[0]
        if len(title_hits) > 1:
            tables = title_hits
    # Title guessing itself can miss (guess_title grabs nearby page text
    # when the real heading isn't where it expects) -- found live on
    # eti2019_impairment_note12.json: the correct table's own title came
    # back as "consolidated statement of profit or loss" (bled in from a
    # neighbouring page element), so the title-match above never fires and
    # row-count alone ties a 6-row Goodwill fragment against this case's
    # 6-row ground truth. Break that tie (and prefer over row-count
    # generally) by how many of the ground truth's own row LABELS actually
    # appear in a candidate -- content the title-guesser can garble but the
    # table's own extracted labels usually still carry correctly.
    gt_labels = {normspace(r[0]).lower() for r in gt_rows if r and r[0]}
    if gt_labels:
        def _label_hits(t):
            cand_labels = {normspace(r[0]).lower() for r in t["rows"] if r and r[0]}
            return len(gt_labels & cand_labels)
        best = max(tables, key=_label_hits)
        if _label_hits(best) > 0:
            return best
    return min(tables, key=lambda t: abs(len(t["rows"]) - len(gt_rows)))


def main():
    cases = sorted(GT_DIR.glob("*.json"))
    if not cases:
        print("No ground-truth cases found in", GT_DIR)
        return 1

    results = []
    for case_path in cases:
        case = json.loads(case_path.read_text(encoding="utf-8"))
        pdf_path = ROOT / case["source_pdf"]
        page0 = case["source_page"] - 1
        if not pdf_path.exists():
            results.append((case_path.name, None, f"PDF not found: {pdf_path.name}"))
            continue
        tables = X.scan([pdf_path], {page0}, min_rows=2, min_cols=2)
        match = best_matching_table(tables, case["rows"], case.get("title"))
        if match is None:
            scored = {"score": 0.0, "rows": [], "note": "extractor found NO table on that page"}
        else:
            scored = score_table(case["rows"], match["rows"])
        results.append((case_path.name, scored))

    print(f"{'case':<38} {'score':>7}  detail")
    print("-" * 70)
    for name, scored in results:
        note = scored.get("note")
        if note:
            print(f"{name:<38} {scored['score']:>7.3f}  {note}")
            continue
        print(f"{name:<38} {scored['score']:>7.3f}")
        for i, row in enumerate(scored["rows"]):
            if row["row_score"] < 0.999:
                label = (row.get("gt") or row.get("extracted") or [""])[0]
                note = row.get("note", "")
                print(f"    row {i}: score={row['row_score']:.3f}  {label!r}  {note}")
                for cells in (row.get("cells") or []):
                    gt_cell, ex_cell = cells
                    if gt_cell != ex_cell:
                        print(f"        ground truth={gt_cell!r}  extracted={ex_cell!r}")

    if results:
        avg = sum(s["score"] for _, s in results) / len(results)
        print("-" * 70)
        print(f"{'AVERAGE':<38} {avg:>7.3f}  over {len(results)} cases (misses count as 0.0)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
