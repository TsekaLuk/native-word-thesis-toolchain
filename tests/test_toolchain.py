from __future__ import annotations

import tempfile
from pathlib import Path

from scripts.quick_validate import CONFIG, make_fixture

from native_word_thesis.ooxml import NS, make_omath, node_text, polish_docx
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


def test_omml_builder_supports_display_formula_shapes() -> None:
    formula = make_omath(
        [
            {"text": "CTR"},
            " = ",
            {"frac": [["1"], ["10"]]},
            {"sum": {"sub": ["o∈", {"text": "Top10"}], "body": [{"rad": [{"text": "rel"}, "(v,o)"]}]}},
        ]
    )

    assert node_text(formula) == "CTR = 110o∈Top10rel(v,o)"
    assert len(formula.xpath(".//m:f", namespaces=NS)) == 1
    assert len(formula.xpath(".//m:nary", namespaces=NS)) == 1
    assert len(formula.xpath(".//m:rad", namespaces=NS)) == 1
