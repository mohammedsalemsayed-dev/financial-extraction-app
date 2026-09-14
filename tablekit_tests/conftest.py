"""Shared pytest fixtures for tablekit_tests/."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import serve  # noqa: E402


@pytest.fixture(autouse=True)
def _isolate_session_dir(monkeypatch, tmp_path):
    """Every test that exercises extract_region / delete_manual / reanalyze
    now triggers serve.py's session-autosave (see `_save_session`), which
    writes real files to disk. Without this, that would land in the actual
    project's uploads/.sessions/ on every test run -- gitignored, so not a
    repo-hygiene problem, but still a test writing outside its own sandbox.
    Autouse + a per-test tmp_path keeps every test hermetic regardless of
    whether it thinks about persistence at all."""
    monkeypatch.setattr(serve, "SESSION_DIR", tmp_path / ".sessions")
