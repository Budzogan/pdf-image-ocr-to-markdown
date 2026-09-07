"""Converter import must not load Docling."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_importing_converter_does_not_load_docling():
    code = (
        "import sys\n"
        "import scan_to_markdown_docling\n"
        "loaded = [name for name in sys.modules if name == 'docling' or name.startswith('docling.')]\n"
        "assert not loaded, loaded\n"
    )
    env = os.environ.copy()
    env["PYTHONPATH"] = str(ROOT)
    result = subprocess.run(
        [sys.executable, "-c", code],
        cwd=ROOT,
        capture_output=True,
        text=True,
        env=env,
    )
    assert result.returncode == 0, result.stderr
