"""
tablekit -- shared pieces of the generic table extractor.

The bulk of the engine still lives in ../extract_all_tables.py; these submodules
hold the parts that are genuinely self-contained and worth testing / reusing on
their own:

    tablekit.config   the CONFIG dict + its TOML override loader
    tablekit.parse    parse_number() and the raw-cell coercion it backs

extract_all_tables.py imports from here, so `import extract_all_tables` keeps
working unchanged.
"""
from .config import CONFIG, load_config_overrides            # noqa: F401
from .parse import parse_number, coerce_cell, normspace       # noqa: F401

__all__ = ["CONFIG", "load_config_overrides",
           "parse_number", "coerce_cell", "normspace"]
