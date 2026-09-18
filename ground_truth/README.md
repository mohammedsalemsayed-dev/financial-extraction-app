# Ground-truth set

Small, hand-verified sample used to regression-test `extract_all_tables.py`'s
output against known-correct values, via `tablekit/scorer.py` /
`score_extraction.py`. Not FinTabNet-scale — deliberately a starting set of
real, previously-audited tables (see project conversation history for how
each was verified: either arithmetic-checked balance sheet/cash flow totals,
or cross-year-consistency-checked note figures read directly from the PDF).

Each `*.json` file: `{"source_pdf": ..., "source_page": <1-indexed>, "title": ...,
"rows": [[label, value_year1, value_year2], ...]}`. `rows` is ground truth —
edit only if you find and fix a transcription error, never to make a score
look better.

To add a new case: hand-transcribe a table's rows directly from the PDF
(never from this tool's own output — that would just test the tool against
itself), save it here, then run:

    python ../score_extraction.py case_name.json <(python -c "...extract & dump rows as JSON...")

See `run_ground_truth_suite.py` (project root) for the automated version that
runs every case here against a live extraction and prints a summary table.
