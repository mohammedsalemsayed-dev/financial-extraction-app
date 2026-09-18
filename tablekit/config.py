"""Every tuned threshold in one place.

Override any key from a TOML file:
  * `extract_all_tables.toml` next to the script, or
  * the file named by $EXTRACT_TABLES_CONFIG.
See `extract_all_tables.toml.example` for the full list with defaults.
"""
from __future__ import annotations

import os
from pathlib import Path

CONFIG = {
    # ---- detection -------------------------------------------------------
    "two_up_page_width": 700,      # a page wider than this is a landscape 2-up spread
    "overlap_dedup": 0.55,         # two boxes overlapping more than this are the same table
    "sliver_max_width": 100,       # a labelless candidate narrower than this (points) is a
                                    # rule-intersection artifact, not a real table -- evictable
                                    # by a better candidate over the same region (see
                                    # find_all_tables); real, single-column-but-legitimate
                                    # tables run wider than this in every real file seen so far
    "recon_head_max_start": 40,    # a statement heading must begin within N chars of line start
    "recon_head_max_len": 95,      # ...and the heading line be no longer than this
    "recon_min_anchor_words": 3,   # a reconstruction needs >= N known statement line-items
    # ---- arithmetic tolerances (statement's own units, usually thousands)
    "tol_pl": 2.0,                 # profit & loss / notes reconciliation
    "tol_bs": 2.0,                 # balance-sheet identities
    "tol_cashflow": 200.0,         # cash flow (rounding drift is wider; also scaled to size)
    "tol_equity": 5.0,             # changes-in-equity column check
    # ---- statement-vs-noise downgrade ----------------------------------
    "prose_row_ratio": 0.30,       # > this fraction of figure-less sentence rows -> not a statement
    "prose_min_words": 7,          # a figure-less label this long reads as prose
    "max_year_gap": 2,             # header years further apart than this are mis-read
    # ---- label / figure health ---------------------------------------
    "health_bad_below": 0.75,      # health score under this is flagged in the UI / audit
    "figure_outlier_sd": 4.0,      # figure > mean + N*sd of its column's log-magnitudes is suspect
    # ---- rendering / scanned-PDF guard ------------------------------
    "page_render_scale": 2.0,      # PDF page -> PNG zoom for the preview
    "scanned_chars_per_page": 90,  # fewer extractable chars/page than this looks scanned
}


def load_config_overrides(announce=print):
    """Merge recognised keys from a TOML override file into CONFIG (in place)."""
    paths = [os.environ.get("EXTRACT_TABLES_CONFIG"),
             str(Path(__file__).resolve().parent.parent / "extract_all_tables.toml")]
    for p in paths:
        if not p or not Path(p).exists():
            continue
        try:
            import tomllib
            with open(p, "rb") as fh:
                data = tomllib.load(fh)
            applied = {k: v for k, v in data.items() if k in CONFIG}
            CONFIG.update(applied)
            if applied and announce:
                announce(f"  (config: {len(applied)} override(s) from {p})")
        except Exception as e:                       # pragma: no cover
            if announce:
                announce(f"  (config: could not read {p}: {e})")
