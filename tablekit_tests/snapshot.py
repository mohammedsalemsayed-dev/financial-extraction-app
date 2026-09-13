"""Regenerate golden.json -- run after a DELIBERATE behaviour change."""
import json, sys
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import extract_all_tables as X
from test_golden import CASE_PAGES, STMT, _prange, _series

snap = []
for f in sorted(CASE_PAGES):
    p = ROOT / f
    if not p.exists():
        print("MISSING", f); continue
    for t in X.scan([p], _prange(CASE_PAGES[f]), 2, 2, warn=lambda *a: None):
        if t["kind"] not in STMT:
            continue
        snap.append({
            "file": f, "page": t["page_label"], "kind": t["kind"],
            "years": t.get("years"), "foots": t.get("foots"),
            "nrows": len(t["rows"]), "ncols": max(len(r) for r in t["rows"]),
            "title": t.get("title"), "series": _series(t),
        })
(Path(__file__).parent / "golden.json").write_text(
    json.dumps(snap, indent=1, ensure_ascii=False), encoding="utf-8")
print(f"wrote {len(snap)} snapshots")
