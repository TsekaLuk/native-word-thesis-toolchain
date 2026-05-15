from __future__ import annotations

import tempfile
from pathlib import Path

from scripts.quick_validate import CONFIG, make_fixture

from native_word_thesis.ooxml import polish_docx
from native_word_thesis.validate import validate_docx


def test_polish_removes_common_conversion_artifacts() -> None:
    with tempfile.TemporaryDirectory(prefix="nwt-test-") as td:
        root = Path(td)
        fixture = root / "fixture.docx"
        output = root / "output.docx"
        make_fixture(fixture)
        polish_docx(fixture, output, CONFIG)
        report = validate_docx(output, CONFIG)

    assert report["ok"]
    assert report["reference_numPr_left"] == 0
    assert report["math_object_count"] >= 1
    assert report["heading_italic_left"] == []
