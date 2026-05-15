from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

from .ooxml import NS, W, node_text, parse_caption, qname


def next_content(children: list[etree._Element], idx: int) -> etree._Element | None:
    for candidate in children[idx + 1 :]:
        if candidate.tag == qname(W, "tbl"):
            return candidate
        if candidate.tag == qname(W, "p") and node_text(candidate):
            return candidate
    return None


def prev_content(children: list[etree._Element], idx: int) -> etree._Element | None:
    for candidate in reversed(children[:idx]):
        if candidate.tag == qname(W, "tbl"):
            return candidate
        if candidate.tag == qname(W, "p") and (node_text(candidate) or candidate.xpath(".//w:drawing|.//w:pict", namespaces=NS)):
            return candidate
    return None


def validate_docx(path: Path, config: dict[str, Any] | None = None) -> dict[str, Any]:
    config = config or {}
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        root = etree.fromstring(zf.read("word/document.xml"))
    body = root.xpath("//w:body", namespaces=NS)[0]
    children = list(body)
    refs = []
    in_refs = False
    heading = config.get("references", {}).get("heading", "参考文献")
    stop_prefixes = tuple(config.get("references", {}).get("stop_prefixes", ["附录", "Appendix"]))
    for child in children:
        if child.tag != qname(W, "p"):
            continue
        text = node_text(child)
        if text == heading:
            in_refs = True
            continue
        if any(text.startswith(prefix) for prefix in stop_prefixes):
            in_refs = False
        if in_refs and text:
            refs.append(child)
    captions = []
    for idx, child in enumerate(children):
        if child.tag != qname(W, "p"):
            continue
        parsed = parse_caption(node_text(child))
        if not parsed:
            continue
        next_node = next_content(children, idx)
        prev_node = prev_content(children, idx)
        real_table = parsed[0] == "表" and next_node is not None and next_node.tag == qname(W, "tbl")
        real_figure = (
            parsed[0] == "图"
            and prev_node is not None
            and bool(prev_node.xpath(".//w:drawing|.//w:pict", namespaces=NS))
        )
        if not (real_table or real_figure):
            continue
        captions.append(
            {
                "text": node_text(child),
                "kind": parsed[0],
                "centered": bool(child.xpath("./w:pPr/w:jc[@w:val='center']", namespaces=NS)),
                "keep_next": bool(child.xpath("./w:pPr/w:keepNext", namespaces=NS)),
                "near_table": real_table,
            }
        )
    heading_italic = []
    pattern = re.compile(r"^(?:\d+(?:\.\d+){2,3}|[A-Z]\.\d+)\s+")
    for p in root.xpath("//w:body/w:p", namespaces=NS):
        text = node_text(p)
        style_vals = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
        style = style_vals[0] if style_vals else ""
        if (style in {"Heading3", "3", "30"} or pattern.match(text)) and p.xpath(".//w:rPr/w:i|.//w:rPr/w:iCs", namespaces=NS):
            heading_italic.append(text)
    report = {
        "zip_ok": bad is None,
        "section_count": len(root.xpath("//w:sectPr", namespaces=NS)),
        "section_header_ref_count": len(root.xpath("//w:sectPr/w:headerReference", namespaces=NS)),
        "section_page_start_count": len(root.xpath("//w:sectPr/w:pgNumType[@w:start]", namespaces=NS)),
        "reference_count": len(refs),
        "reference_numPr_left": sum(len(p.xpath("./w:pPr/w:numPr", namespaces=NS)) for p in refs),
        "math_object_count": len(root.xpath("//m:oMath", namespaces=NS)),
        "table_count": len(root.xpath("//w:tbl", namespaces=NS)),
        "centered_table_count": len(root.xpath("//w:tbl/w:tblPr/w:jc[@w:val='center']", namespaces=NS)),
        "caption_count": len(captions),
        "uncentered_captions": [c["text"] for c in captions if not c["centered"]],
        "floating_table_captions": [c["text"] for c in captions if c["kind"] == "表" and not c["keep_next"]],
        "heading_italic_left": heading_italic,
    }
    report["ok"] = (
        report["zip_ok"]
        and report["reference_numPr_left"] == 0
        and not report["uncentered_captions"]
        and not report["floating_table_captions"]
        and not report["heading_italic_left"]
    )
    return report


def dump_report(report: dict[str, Any], output: Path | None = None) -> str:
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if output:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(text + "\n", encoding="utf-8")
    return text
