"""Cell/row-level accuracy scorer for regression-testing our own extractions
against hand-transcribed ground truth.

Deliberately NOT a TEDS/FinTabNet-style structural scorer -- this tool emits
flat row/column grids with no rowspan/colspan concept, so a structural score
would mostly measure that gap rather than whether the numbers are right. This
scores CONTENT: did the extractor find the right values, in the right cells,
in roughly the right order. Method: cell-level similarity (numeric cells
compared via parse_number for exact-value equality; text cells via a
normalized Levenshtein ratio), rolled up through a row alignment done with
global sequence alignment (Needleman-Wunsch) with *free end gaps* -- so a
ground-truth table that's a sub-range of what got extracted (or vice versa,
e.g. one extra header/footnote row at either end) isn't penalized for rows
outside the overlap, while a genuinely missing/extra/shifted row in the
middle still costs, since that's a real defect.
"""
from __future__ import annotations

from .parse import parse_number, normspace

GAP_PENALTY = 1.0     # cost of aligning a row/cell against a gap (internal only)
MISMATCH_FLOOR = 0.0  # worst-case similarity for two rows that share nothing


def _levenshtein(a, b):
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i] + [0] * len(b)
        for j, cb in enumerate(b, 1):
            cost = 0 if ca == cb else 1
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
        prev = cur
    return prev[-1]


def cell_similarity(a, b):
    """1.0 = identical value, 0.0 = no relation. Numbers must match exactly
    (a financial figure that's off by even one digit is simply wrong, no
    partial credit); text is scored by normalized edit distance."""
    an, bn = parse_number(a), parse_number(b)
    if an is not None or bn is not None:
        if an is None or bn is None:
            return MISMATCH_FLOOR
        return 1.0 if float(an) == float(bn) else MISMATCH_FLOOR
    sa = normspace(a).lower() if a is not None else ""
    sb = normspace(b).lower() if b is not None else ""
    if sa == sb:
        return 1.0
    if not sa and not sb:
        return 1.0
    if not sa or not sb:
        return MISMATCH_FLOOR
    dist = _levenshtein(sa, sb)
    return max(0.0, 1.0 - dist / max(len(sa), len(sb)))


def _align(seq_a, seq_b, sim_fn, gap_penalty=GAP_PENALTY):
    """Needleman-Wunsch global alignment with free end gaps. Returns
    (score_0_to_1, pairs) where pairs is a list of (a_item_or_None,
    b_item_or_None). score is the mean similarity over MATCHED pairs only
    (gaps at either end are free and excluded; internal gaps count as 0
    similarity, dragging the mean down as they should)."""
    n, m = len(seq_a), len(seq_b)
    if n == 0 and m == 0:
        return 1.0, []
    # dp[i][j] = best score aligning seq_a[:i] with seq_b[:j]; free end gaps
    # means the first row/column costs 0, not i*gap_penalty / j*gap_penalty.
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    back = [[None] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        back[i][0] = 'up'
    for j in range(1, m + 1):
        back[0][j] = 'left'
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            diag = dp[i - 1][j - 1] + sim_fn(seq_a[i - 1], seq_b[j - 1])
            up = dp[i - 1][j] - (0 if j == m else gap_penalty)
            left = dp[i][j - 1] - (0 if i == n else gap_penalty)
            best = max(diag, up, left)
            dp[i][j] = best
            back[i][j] = 'diag' if best == diag else ('up' if best == up else 'left')
    pairs = []
    i, j = n, m
    while i > 0 or j > 0:
        move = back[i][j]
        if move == 'diag':
            pairs.append((seq_a[i - 1], seq_b[j - 1]))
            i, j = i - 1, j - 1
        elif move == 'up':
            pairs.append((seq_a[i - 1], None))
            i -= 1
        else:
            pairs.append((None, seq_b[j - 1]))
            j -= 1
    pairs.reverse()
    matched = [(a, b) for a, b in pairs if a is not None and b is not None]
    internal_gaps = sum(1 for a, b in pairs if a is None or b is None) - \
        _leading_trailing_gap_count(pairs)
    if not matched and internal_gaps == 0:
        return 1.0, pairs
    total_scored = len(matched) + internal_gaps
    if total_scored == 0:
        return 1.0, pairs
    score = sum(sim_fn(a, b) for a, b in matched) / total_scored
    return score, pairs


def _leading_trailing_gap_count(pairs):
    n = 0
    for a, b in pairs:
        if a is None or b is None:
            n += 1
        else:
            break
    for a, b in reversed(pairs):
        if a is None or b is None:
            n += 1
        else:
            break
    return n


def score_row(gt_row, ex_row):
    """Align two rows' cells (handles an extractor that dropped/added a
    column) and return (score_0_to_1, cell_pairs)."""
    return _align(list(gt_row), list(ex_row), cell_similarity)


def score_table(gt_rows, ex_rows):
    """Align two tables' rows, each row-pair itself scored by score_row.
    Returns a dict: overall score, per-row breakdown, and the alignment
    (which ground-truth row matched which extracted row, if any) -- useful
    for pinpointing exactly which row/cell went wrong, not just a number."""
    def row_sim(a, b):
        s, _ = score_row(a, b)
        return s

    score, pairs = _align(list(gt_rows), list(ex_rows), row_sim)
    breakdown = []
    for gt_row, ex_row in pairs:
        if gt_row is None:
            breakdown.append({"gt": None, "extracted": ex_row, "row_score": 0.0,
                               "note": "extra row (not in ground truth)"})
        elif ex_row is None:
            breakdown.append({"gt": gt_row, "extracted": None, "row_score": 0.0,
                               "note": "missing row (in ground truth, not extracted)"})
        else:
            row_score, cell_pairs = score_row(gt_row, ex_row)
            breakdown.append({"gt": gt_row, "extracted": ex_row, "row_score": row_score,
                               "cells": cell_pairs})
    return {"score": score, "rows": breakdown}
