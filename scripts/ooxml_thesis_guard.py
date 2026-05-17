#!/usr/bin/env python3
"""Audit and repair reusable OOXML thesis-format invariants.

This utility intentionally works below python-docx. Thesis formatting bugs that
look like "one pixel off" in WPS/Word often live in raw OOXML: false bold/italic
toggle nodes, inherited header indents, style-only header borders, or captions
centered inside a first-line indent. The config controls school-specific
granularity while the checks stay reusable across projects.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Any

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"w": W, "m": M, "wp": WP}
EMU_PER_INCH = 914400
MATH_OPERATOR_TEXT_RE = re.compile(r"^[\s=+\-−×÷*/≤≥<>:|,;(){}\[\]∈∩∪]+$")
CITATION_SUPERSCRIPT_RE = re.compile(r"^\[?[\d\s,，;；\-–—]+]?$")


DEFAULT_CONFIG: dict[str, Any] = {
    "heading_styles": ["Heading1", "Heading2", "Heading3"],
    "forbid_heading4": True,
    "forbid_heading_bold": True,
    "forbid_heading_italic": True,
    "strip_heading_toggle_noise": True,
    "strip_caption_toggle_noise": True,
    "style_gallery": {
        "required_names": {},
        "visible_qformat_styles": [],
        "hidden_qformat_styles": [],
    },
    "conversion_styles": {
        "forbid_paragraph_styles": ["Compact", "FirstParagraph", "CaptionedFigure"],
        "body_style": "BodyText",
        "table_body_style": "TableBody",
    },
    "caption_pattern": r"^(图|表)\s*\d+[-.]\d+\s+",
    "exclude_list_entries_with_tabs": True,
    "require_caption_center": True,
    "require_caption_zero_indent": True,
    "require_drawing_center": True,
    "require_drawing_zero_indent": True,
    "drawing_max_width_in": None,
    "drawing_max_height_in": None,
    "header": {
        "required_text_contains": "",
        "require_zero_indent": True,
        "require_bottom_border": True,
        "bottom_border_color": "7F7F7F",
        "bottom_border_size": "6",
        "bottom_border_space": "1",
        "require_nbsp_in_page_label": False,
        "required_spacer_len": None,
        "required_right_tab_pos": None,
    },
    "references": {
        "headings_compact": ["参考文献"],
        "stop_heading_prefixes": ["附录", "Appendix"],
        "forbid_decimal_plus_bracket_numbering": True,
        "forbid_word_wrap": False,
        "forbid_hard_line_breaks": False,
        "max_text_runs_per_entry": None,
    },
    "citations": {
        "require_superscript_brackets": True,
        "scan_before_reference_heading": True,
        "reference_headings_compact": ["参考文献"],
    },
    "front_matter": {
        "scan_first_paragraphs": 90,
        "forbid_empty_break_paragraphs": True,
    },
    "field_codes": {
        "forbid_instr_text_patterns": [],
        "forbid_visible_text_patterns": [],
    },
    "latin_fonts": {
        "required_ascii_hansi": None,
        "require_direct_on_latin_runs": False,
    },
    "math": {
        "min_omml_count": None,
        "min_subscript_count": None,
        "forbid_simple_numeric_omml": False,
        "required_math_font": None,
        "allowed_math_run_fonts": [],
        "upright_text_math_font": None,
        "operator_math_font": None,
        "forbid_spaced_operator_runs": False,
        "forbidden_math_fonts": [],
        "require_direct_math_run_font": False,
        "required_formula_fragments": [],
    },
}


def qn(tag: str) -> str:
    prefix, name = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{name}"


def w_el(tag: str) -> etree._Element:
    return etree.Element(f"{{{W}}}{tag}")


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_config(path: Path | None) -> dict[str, Any]:
    if path is None:
        return deepcopy(DEFAULT_CONFIG)
    return deep_merge(DEFAULT_CONFIG, json.loads(path.read_text(encoding="utf-8")))


def paragraph_text(p: etree._Element) -> str:
    return "".join(p.xpath(".//w:t/text() | .//m:t/text()", namespaces=NS))


def visible_or_field_paragraph_text(p: etree._Element) -> str:
    return "".join(p.xpath(".//w:t/text() | .//w:instrText/text() | .//m:t/text()", namespaces=NS))


def ancestor_paragraph(node: etree._Element) -> etree._Element | None:
    current: etree._Element | None = node
    while current is not None:
        if current.tag == qn("w:p"):
            return current
        current = current.getparent()
    return None


def compact_text(text: str) -> str:
    return re.sub(r"\s+", "", text.replace("\u00a0", ""))


def math_run_text(run: etree._Element) -> str:
    return "".join(run.xpath(".//m:t/text()", namespaces=NS))


def is_operator_math_text(text: str) -> bool:
    return bool(MATH_OPERATOR_TEXT_RE.fullmatch(text)) and any(not ch.isspace() for ch in text)


def math_run_prefers_text_font(run: etree._Element) -> bool:
    text = math_run_text(run)
    return bool(run.xpath("./m:rPr/m:nor", namespaces=NS)) or bool(re.search(r"[A-Za-z]{2,}|@", text))


def style_id(p: etree._Element) -> str:
    node = p.find("./w:pPr/w:pStyle", namespaces=NS)
    return node.get(qn("w:val"), "") if node is not None else ""


def has_ancestor(node: etree._Element, tag: str) -> bool:
    current = node.getparent()
    target = qn(tag)
    while current is not None:
        if current.tag == target:
            return True
        current = current.getparent()
    return False


def set_paragraph_style(p: etree._Element, sid: str) -> None:
    p_pr = get_or_add(p, "pPr")
    p_style = p_pr.find("w:pStyle", namespaces=NS)
    if p_style is None:
        p_style = etree.Element(f"{{{W}}}pStyle")
        p_pr.insert(0, p_style)
    p_style.set(qn("w:val"), sid)


def superscript_groups(p: etree._Element) -> list[str]:
    groups: list[str] = []
    current: list[str] = []
    for run in p.xpath("./w:r", namespaces=NS):
        text = "".join(run.xpath(".//w:t/text()", namespaces=NS))
        vert = run.find("./w:rPr/w:vertAlign", namespaces=NS)
        is_superscript = vert is not None and vert.get(qn("w:val")) == "superscript"
        if is_superscript:
            current.append(text)
        elif current:
            groups.append("".join(current))
            current = []
    if current:
        groups.append("".join(current))
    return groups


def is_citation_like_superscript(text: str) -> bool:
    stripped = text.strip()
    return bool(stripped and any(ch.isdigit() for ch in stripped) and CITATION_SUPERSCRIPT_RE.fullmatch(stripped))


def jc_val(p: etree._Element) -> str | None:
    node = p.find("./w:pPr/w:jc", namespaces=NS)
    return node.get(qn("w:val")) if node is not None else None


def ind_attrs(p: etree._Element) -> dict[str, str]:
    node = p.find("./w:pPr/w:ind", namespaces=NS)
    if node is None:
        return {}
    return {key.split("}", 1)[1]: value for key, value in node.attrib.items()}


def get_or_add(parent: etree._Element, tag: str) -> etree._Element:
    node = parent.find(f"w:{tag}", namespaces=NS)
    if node is None:
        node = etree.SubElement(parent, f"{{{W}}}{tag}")
    return node


def get_or_add_ppr(p: etree._Element) -> etree._Element:
    p_pr = p.find("w:pPr", namespaces=NS)
    if p_pr is None:
        p_pr = w_el("pPr")
        p.insert(0, p_pr)
    return p_pr


def set_style_child_val(style: etree._Element, child: str, value: str) -> None:
    node = style.find(f"w:{child}", namespaces=NS)
    if node is None:
        node = etree.SubElement(style, f"{{{W}}}{child}")
    node.set(qn("w:val"), value)


def remove_style_child(style: etree._Element, child: str) -> None:
    for node in style.xpath(f"./w:{child}", namespaces=NS):
        style.remove(node)


def ensure_style_child(style: etree._Element, child: str) -> None:
    if style.find(f"w:{child}", namespaces=NS) is None:
        etree.SubElement(style, f"{{{W}}}{child}")


def set_style_gallery_visibility(style: etree._Element, visible: bool) -> None:
    if visible:
        ensure_style_child(style, "qFormat")
        remove_style_child(style, "semiHidden")
        remove_style_child(style, "hidden")
    else:
        remove_style_child(style, "qFormat")
        ensure_style_child(style, "semiHidden")
        ensure_style_child(style, "unhideWhenUsed")


def ensure_jc(p: etree._Element, value: str) -> None:
    p_pr = get_or_add_ppr(p)
    jc = get_or_add(p_pr, "jc")
    jc.set(qn("w:val"), value)


def ensure_zero_indent(p: etree._Element) -> None:
    p_pr = get_or_add_ppr(p)
    ind = get_or_add(p_pr, "ind")
    for attr in ["left", "right", "firstLine", "hanging"]:
        ind.set(qn(f"w:{attr}"), "0")


def toggle_effective(r_pr: etree._Element | None, tag: str) -> bool:
    if r_pr is None:
        return False
    for node in r_pr.xpath(f"./w:{tag}", namespaces=NS):
        value = node.get(qn("w:val"))
        if value not in {"0", "false", "False", "off", "none"}:
            return True
    return False


def toggle_any(r_pr: etree._Element | None, tag: str) -> bool:
    return bool(r_pr is not None and r_pr.xpath(f"./w:{tag}", namespaces=NS))


def remove_toggle_nodes(r_pr: etree._Element, tags: list[str]) -> None:
    for tag in tags:
        for node in r_pr.xpath(f"./w:{tag}", namespaces=NS):
            r_pr.remove(node)


def paragraph_has_drawing(p: etree._Element) -> bool:
    return bool(p.xpath(".//wp:inline|.//wp:anchor", namespaces=NS))


def paragraph_max_drawing_inches(p: etree._Element) -> tuple[float, float]:
    max_w = 0.0
    max_h = 0.0
    for ext in p.xpath(".//wp:extent", namespaces=NS):
        max_w = max(max_w, int(ext.get("cx", "0")) / EMU_PER_INCH)
        max_h = max(max_h, int(ext.get("cy", "0")) / EMU_PER_INCH)
    return max_w, max_h


def ensure_header_bottom_border(p: etree._Element, cfg: dict[str, Any]) -> None:
    p_pr = get_or_add_ppr(p)
    p_bdr = get_or_add(p_pr, "pBdr")
    bottom = get_or_add(p_bdr, "bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:color"), cfg.get("bottom_border_color", "7F7F7F"))
    bottom.set(qn("w:sz"), str(cfg.get("bottom_border_size", "6")))
    bottom.set(qn("w:space"), str(cfg.get("bottom_border_space", "1")))


def iter_docx_xml(root_dir: Path) -> tuple[etree._ElementTree, etree._Element, etree._ElementTree, etree._Element, dict[str, etree._ElementTree]]:
    document_tree = etree.parse(str(root_dir / "word/document.xml"))
    body = document_tree.xpath("//w:body", namespaces=NS)[0]
    styles_tree = etree.parse(str(root_dir / "word/styles.xml"))
    styles_root = styles_tree.getroot()
    headers: dict[str, etree._ElementTree] = {}
    for path in sorted((root_dir / "word").glob("header*.xml")):
        headers[path.name] = etree.parse(str(path))
    return document_tree, body, styles_tree, styles_root, headers


def audit_unpacked(root_dir: Path, cfg: dict[str, Any], fix: bool = False) -> dict[str, Any]:
    document_tree, body, styles_tree, styles_root, headers = iter_docx_xml(root_dir)
    paragraphs = body.xpath("./w:p", namespaces=NS)
    warnings: list[dict[str, Any]] = []
    fixes: list[str] = []

    heading_styles = set(cfg["heading_styles"])
    heading_counts: dict[str, int] = {style: 0 for style in [*cfg["heading_styles"], "Heading4"]}
    heading_toggle_noise: list[dict[str, Any]] = []
    heading_effective_toggles: list[dict[str, Any]] = []
    for idx, p in enumerate(paragraphs):
        sid = style_id(p)
        if sid not in heading_styles and sid != "Heading4":
            continue
        heading_counts[sid] = heading_counts.get(sid, 0) + 1
        text = paragraph_text(p)
        if sid == "Heading4" and cfg.get("forbid_heading4", True):
            warnings.append({"code": "heading4_present", "paragraph": idx, "text": text[:120]})
        effective = False
        noisy = False
        for r_pr in p.xpath(".//w:rPr|./w:pPr/w:rPr", namespaces=NS):
            bold = toggle_effective(r_pr, "b") or toggle_effective(r_pr, "bCs")
            italic = toggle_effective(r_pr, "i") or toggle_effective(r_pr, "iCs")
            any_toggle = any(toggle_any(r_pr, tag) for tag in ["b", "bCs", "i", "iCs"])
            if bold and cfg.get("forbid_heading_bold", True):
                effective = True
            if italic and cfg.get("forbid_heading_italic", True):
                effective = True
            if any_toggle:
                noisy = True
                if fix and cfg.get("strip_heading_toggle_noise", True):
                    remove_toggle_nodes(r_pr, ["b", "bCs", "i", "iCs"])
        if effective:
            heading_effective_toggles.append({"paragraph": idx, "style": sid, "text": text[:120]})
        if noisy:
            heading_toggle_noise.append({"paragraph": idx, "style": sid, "text": text[:120]})
            if fix and cfg.get("strip_heading_toggle_noise", True):
                fixes.append(f"cleared heading toggles at paragraph {idx}")

    for style in styles_root.xpath('.//w:style[@w:styleId="Heading1" or @w:styleId="Heading2" or @w:styleId="Heading3"]', namespaces=NS):
        for r_pr in style.xpath("./w:rPr", namespaces=NS):
            if any(toggle_any(r_pr, tag) for tag in ["b", "bCs", "i", "iCs"]):
                if fix and cfg.get("strip_heading_toggle_noise", True):
                    remove_toggle_nodes(r_pr, ["b", "bCs", "i", "iCs"])
                    fixes.append(f"cleared {style.get(qn('w:styleId'))} style toggles")
                else:
                    warnings.append({"code": "heading_style_toggle_noise", "style": style.get(qn("w:styleId"))})

    if cfg.get("strip_caption_toggle_noise", True):
        for style in styles_root.xpath('.//w:style[@w:styleId="Caption"]', namespaces=NS):
            for r_pr in style.xpath("./w:rPr", namespaces=NS):
                if any(toggle_any(r_pr, tag) for tag in ["b", "bCs", "i", "iCs"]):
                    if fix:
                        remove_toggle_nodes(r_pr, ["b", "bCs", "i", "iCs"])
                        fixes.append("cleared Caption style toggles")
                    else:
                        warnings.append({"code": "caption_style_toggle_noise"})

    citations_cfg = cfg.get("citations", {})
    citation_superscript_violations: list[dict[str, Any]] = []
    ref_heading_compact = set(citations_cfg.get("reference_headings_compact", ["参考文献"]))
    citation_scan_stop = len(paragraphs)
    if citations_cfg.get("scan_before_reference_heading", True):
        for idx, p in enumerate(paragraphs):
            if compact_text(paragraph_text(p)) in ref_heading_compact:
                citation_scan_stop = idx
                break
    if citations_cfg.get("require_superscript_brackets", True):
        for idx, p in enumerate(paragraphs[:citation_scan_stop]):
            para_text = paragraph_text(p)
            for superscript in superscript_groups(p):
                stripped = superscript.strip()
                if not is_citation_like_superscript(stripped):
                    continue
                if stripped.startswith("[") and stripped.endswith("]"):
                    continue
                violation = {
                    "paragraph": idx,
                    "superscript": stripped,
                    "text": para_text[:120],
                }
                citation_superscript_violations.append(violation)
                warnings.append({"code": "citation_superscript_missing_brackets", **violation})

    style_gallery_cfg = cfg.get("style_gallery", {})
    style_gallery_checks: list[dict[str, Any]] = []
    for sid_required, expected_name in style_gallery_cfg.get("required_names", {}).items():
        nodes = styles_root.xpath(f'.//w:style[@w:styleId="{sid_required}"]', namespaces=NS)
        if not nodes:
            warnings.append({"code": "style_missing", "styleId": sid_required, "expected_name": expected_name})
            continue
        style = nodes[0]
        name_node = style.find("w:name", namespaces=NS)
        actual_name = name_node.get(qn("w:val")) if name_node is not None else None
        item = {"styleId": sid_required, "expected": expected_name, "actual": actual_name}
        style_gallery_checks.append(item)
        if actual_name != expected_name:
            warnings.append({"code": "style_name_mismatch", **item})
            if fix:
                set_style_child_val(style, "name", expected_name)
                fixes.append(f"renamed style {sid_required} to {expected_name}")

    for sid_visible in style_gallery_cfg.get("visible_qformat_styles", []):
        nodes = styles_root.xpath(f'.//w:style[@w:styleId="{sid_visible}"]', namespaces=NS)
        if not nodes:
            continue
        style = nodes[0]
        is_visible = style.find("w:qFormat", namespaces=NS) is not None and style.find("w:semiHidden", namespaces=NS) is None
        if not is_visible:
            warnings.append({"code": "style_should_be_visible_in_gallery", "styleId": sid_visible})
            if fix:
                set_style_gallery_visibility(style, True)
                fixes.append(f"made style {sid_visible} visible in gallery")

    for sid_hidden in style_gallery_cfg.get("hidden_qformat_styles", []):
        nodes = styles_root.xpath(f'.//w:style[@w:styleId="{sid_hidden}"]', namespaces=NS)
        if not nodes:
            continue
        style = nodes[0]
        is_hidden = style.find("w:qFormat", namespaces=NS) is None and style.find("w:semiHidden", namespaces=NS) is not None
        if not is_hidden:
            warnings.append({"code": "style_should_be_hidden_from_gallery", "styleId": sid_hidden})
            if fix:
                set_style_gallery_visibility(style, False)
                fixes.append(f"hid style {sid_hidden} from gallery")

    conversion_cfg = cfg.get("conversion_styles", {})
    forbidden_conversion_styles = set(conversion_cfg.get("forbid_paragraph_styles", []))
    conversion_style_hits: list[dict[str, Any]] = []
    if forbidden_conversion_styles:
        for idx, p in enumerate(paragraphs):
            sid = style_id(p)
            if sid not in forbidden_conversion_styles:
                continue
            text = paragraph_text(p)
            hit = {
                "paragraph": idx,
                "styleId": sid,
                "inside_table": has_ancestor(p, "w:tbl"),
                "text": text[:120],
            }
            conversion_style_hits.append(hit)
            warnings.append({"code": "conversion_style_residue", **hit})
            if fix:
                replacement = conversion_cfg.get("table_body_style") if hit["inside_table"] else conversion_cfg.get("body_style")
                if replacement:
                    set_paragraph_style(p, replacement)
                    fixes.append(f"mapped paragraph {idx} style {sid} to {replacement}")

    caption_re = re.compile(cfg["caption_pattern"])
    caption_count = 0
    for idx, p in enumerate(paragraphs):
        text = paragraph_text(p)
        if cfg.get("exclude_list_entries_with_tabs", True) and "\t" in text:
            continue
        if not caption_re.match(text.strip()):
            continue
        caption_count += 1
        bad_align = cfg.get("require_caption_center", True) and jc_val(p) != "center"
        bad_indent = cfg.get("require_caption_zero_indent", True) and any(ind_attrs(p).get(key) not in {None, "0"} for key in ["left", "right", "firstLine", "hanging"])
        if bad_align or bad_indent:
            warnings.append({"code": "caption_alignment_or_indent", "paragraph": idx, "text": text[:120], "jc": jc_val(p), "ind": ind_attrs(p)})
            if fix:
                if cfg.get("require_caption_center", True):
                    ensure_jc(p, "center")
                if cfg.get("require_caption_zero_indent", True):
                    ensure_zero_indent(p)
                fixes.append(f"normalized caption paragraph {idx}")

    drawing_count = 0
    max_drawing_width = 0.0
    max_drawing_height = 0.0
    for idx, p in enumerate(paragraphs):
        if not paragraph_has_drawing(p):
            continue
        drawing_count += 1
        width, height = paragraph_max_drawing_inches(p)
        max_drawing_width = max(max_drawing_width, width)
        max_drawing_height = max(max_drawing_height, height)
        max_w_cfg = cfg.get("drawing_max_width_in")
        max_h_cfg = cfg.get("drawing_max_height_in")
        if max_w_cfg and width > float(max_w_cfg) + 0.01:
            warnings.append({"code": "drawing_too_wide", "paragraph": idx, "width_in": round(width, 3)})
        if max_h_cfg and height > float(max_h_cfg) + 0.01:
            warnings.append({"code": "drawing_too_tall", "paragraph": idx, "height_in": round(height, 3)})
        bad_align = cfg.get("require_drawing_center", True) and jc_val(p) != "center"
        bad_indent = cfg.get("require_drawing_zero_indent", True) and any(ind_attrs(p).get(key) not in {None, "0"} for key in ["left", "right", "firstLine", "hanging"])
        if bad_align or bad_indent:
            warnings.append({"code": "drawing_alignment_or_indent", "paragraph": idx, "jc": jc_val(p), "ind": ind_attrs(p)})
            if fix:
                if cfg.get("require_drawing_center", True):
                    ensure_jc(p, "center")
                if cfg.get("require_drawing_zero_indent", True):
                    ensure_zero_indent(p)
                fixes.append(f"normalized drawing paragraph {idx}")

    header_cfg = cfg.get("header", {})
    header_checks: list[dict[str, Any]] = []
    for name, tree in headers.items():
        for p in tree.xpath("//w:p", namespaces=NS):
            text = paragraph_text(p)
            required = header_cfg.get("required_text_contains") or ""
            if required and required not in text:
                continue
            if not required and not text.strip():
                continue
            check = {
                "file": name,
                "text": text,
                "jc": jc_val(p),
                "ind": ind_attrs(p),
                "has_bottom_border": bool(p.xpath('./w:pPr/w:pBdr/w:bottom[@w:val="single"]', namespaces=NS)),
                "has_nbsp": "\u00a0" in text,
                "spacer_len": max([len(m.group(1)) for m in re.finditer(r"( +)", text)] or [0]),
                "right_tab_positions": [
                    tab.get(qn("w:pos"))
                    for tab in p.xpath('./w:pPr/w:tabs/w:tab[@w:val="right"]', namespaces=NS)
                ],
            }
            header_checks.append(check)
            if header_cfg.get("require_zero_indent", True) and any(check["ind"].get(key) not in {None, "0"} for key in ["left", "right", "firstLine", "hanging"]):
                warnings.append({"code": "header_indent", **check})
                if fix:
                    ensure_zero_indent(p)
                    fixes.append(f"zeroed header indent in {name}")
            if header_cfg.get("require_bottom_border", True) and not check["has_bottom_border"]:
                warnings.append({"code": "header_missing_bottom_border", **check})
                if fix:
                    ensure_header_bottom_border(p, header_cfg)
                    fixes.append(f"added header bottom border in {name}")
            if header_cfg.get("require_nbsp_in_page_label", False) and "页" in text and "\u00a0" not in text:
                warnings.append({"code": "header_page_label_can_break", **check})
            required_spacer = header_cfg.get("required_spacer_len")
            if required_spacer is not None and check["spacer_len"] != int(required_spacer):
                warnings.append({"code": "header_spacer_len_mismatch", "expected": required_spacer, **check})
            required_right_tab = header_cfg.get("required_right_tab_pos")
            if required_right_tab is not None and str(required_right_tab) not in check["right_tab_positions"]:
                warnings.append({"code": "header_right_tab_missing", "expected": str(required_right_tab), **check})

    references_cfg = cfg.get("references", {})
    ref_start: int | None = None
    ref_headings = {compact_text(item) for item in references_cfg.get("headings_compact", ["参考文献"])}
    for idx, p in enumerate(paragraphs):
        if compact_text(paragraph_text(p)) in ref_headings:
            ref_start = idx
            break
    duplicate_reference_numbers = []
    reference_word_wrap_hits: list[dict[str, Any]] = []
    reference_hard_break_hits: list[dict[str, Any]] = []
    reference_text_run_count_violations: list[dict[str, Any]] = []
    reference_paragraphs: list[tuple[int, etree._Element, str]] = []
    if ref_start is not None:
        stop_prefixes = tuple(compact_text(item) for item in references_cfg.get("stop_heading_prefixes", []))
        for idx, p in enumerate(paragraphs[ref_start + 1 :], start=ref_start + 1):
            text = paragraph_text(p).strip()
            compact = compact_text(text)
            if text and (style_id(p).startswith("Heading") or any(compact.startswith(prefix) for prefix in stop_prefixes)):
                break
            if not text:
                continue
            reference_paragraphs.append((idx, p, text))

    if references_cfg.get("forbid_decimal_plus_bracket_numbering", True):
        for idx, _p, text in reference_paragraphs:
            if re.match(r"^\d+\.\s*\[\d+\]", text):
                duplicate_reference_numbers.append({"paragraph": idx, "text": text[:160]})
    for item in duplicate_reference_numbers:
        warnings.append({"code": "reference_double_numbering", **item})

    max_reference_text_runs = references_cfg.get("max_text_runs_per_entry")
    for idx, p, text in reference_paragraphs:
        if references_cfg.get("forbid_word_wrap", False) and p.xpath("./w:pPr/w:wordWrap", namespaces=NS):
            item = {"paragraph": idx, "text": text[:160]}
            reference_word_wrap_hits.append(item)
            warnings.append({"code": "reference_word_wrap_enabled", **item})
        if references_cfg.get("forbid_hard_line_breaks", False) and p.xpath(".//w:br[not(@w:type) or @w:type='textWrapping']", namespaces=NS):
            item = {"paragraph": idx, "text": text[:160]}
            reference_hard_break_hits.append(item)
            warnings.append({"code": "reference_hard_line_break", **item})
        if max_reference_text_runs is not None:
            text_runs = [
                "".join(run.xpath("./w:t/text()", namespaces=NS))
                for run in p.xpath(".//w:r[not(ancestor::m:oMath)]", namespaces=NS)
            ]
            text_runs = [item for item in text_runs if item]
            if len(text_runs) > int(max_reference_text_runs):
                item = {"paragraph": idx, "text_run_count": len(text_runs), "text": text[:160]}
                reference_text_run_count_violations.append(item)
                warnings.append({"code": "reference_text_run_count", "expected_max": int(max_reference_text_runs), **item})

    front_cfg = cfg.get("front_matter", {})
    front_empty_breaks = []
    if front_cfg.get("forbid_empty_break_paragraphs", True):
        limit = int(front_cfg.get("scan_first_paragraphs", 90))
        for idx, p in enumerate(paragraphs[:limit]):
            if paragraph_text(p).strip():
                continue
            if p.xpath(".//w:br|./w:pPr/w:sectPr", namespaces=NS):
                item = {"paragraph": idx, "style": style_id(p), "jc": jc_val(p), "ind": ind_attrs(p)}
                front_empty_breaks.append(item)
                warnings.append({"code": "front_empty_break_paragraph", **item})

    field_cfg = cfg.get("field_codes", {})
    forbidden_field_code_hits: list[dict[str, Any]] = []
    for source, xpath_expr, patterns in [
        ("instrText", ".//w:instrText", field_cfg.get("forbid_instr_text_patterns", [])),
        ("visibleText", ".//w:t", field_cfg.get("forbid_visible_text_patterns", [])),
    ]:
        for node in body.xpath(xpath_expr, namespaces=NS):
            node_text = node.text or ""
            if not node_text:
                continue
            for pattern in patterns:
                if not re.search(pattern, node_text):
                    continue
                owner = ancestor_paragraph(node)
                item = {
                    "source": source,
                    "pattern": pattern,
                    "text": node_text[:160],
                    "paragraph_text": visible_or_field_paragraph_text(owner)[:220] if owner is not None else node_text[:220],
                }
                forbidden_field_code_hits.append(item)
                warnings.append({"code": "forbidden_field_code_text", **item})
                break

    latin_cfg = cfg.get("latin_fonts", {})
    required_latin = latin_cfg.get("required_ascii_hansi")
    require_direct_latin = bool(latin_cfg.get("require_direct_on_latin_runs", False))
    latin_font_violations: list[dict[str, Any]] = []
    if required_latin:
        for r in body.xpath(".//w:r[w:t]", namespaces=NS):
            text = "".join(r.xpath("./w:t/text()", namespaces=NS))
            if not text or not re.search(r"[A-Za-z]", text):
                continue
            r_fonts = r.find("./w:rPr/w:rFonts", namespaces=NS)
            if r_fonts is None:
                if not require_direct_latin:
                    continue
                item = {"text": text[:120], "ascii": None, "hAnsi": None}
                latin_font_violations.append(item)
                warnings.append({"code": "latin_font_not_direct", **item})
                continue
            ascii_font = r_fonts.get(qn("w:ascii"))
            hansi_font = r_fonts.get(qn("w:hAnsi"))
            if ascii_font != required_latin or hansi_font != required_latin:
                item = {"text": text[:120], "ascii": ascii_font, "hAnsi": hansi_font, "expected": required_latin}
                latin_font_violations.append(item)
                warnings.append({"code": "latin_font_mismatch", **item})

    math_cfg = cfg.get("math", {})
    settings_root: etree._Element | None = None
    settings_path = root_dir / "word/settings.xml"
    if settings_path.exists():
        settings_root = etree.parse(str(settings_path)).getroot()
    configured_math_fonts = settings_root.xpath(".//m:mathFont/@m:val", namespaces=NS) if settings_root is not None else []
    required_math_font = math_cfg.get("required_math_font")
    forbidden_math_fonts = set(math_cfg.get("forbidden_math_fonts") or [])
    if required_math_font and required_math_font not in configured_math_fonts:
        warnings.append(
            {
                "code": "math_font_missing_or_mismatch",
                "expected": required_math_font,
                "actual": configured_math_fonts,
            }
        )
    forbidden_math_font_hits: list[dict[str, Any]] = []
    for font in configured_math_fonts:
        if font in forbidden_math_fonts:
            item = {"location": "settings.xml", "font": font}
            forbidden_math_font_hits.append(item)
            warnings.append({"code": "forbidden_math_font", **item})
    allowed_math_run_fonts = set(math_cfg.get("allowed_math_run_fonts") or [])
    if not allowed_math_run_fonts and required_math_font:
        allowed_math_run_fonts = {required_math_font}
    upright_text_math_font = math_cfg.get("upright_text_math_font")
    operator_math_font = math_cfg.get("operator_math_font")
    forbid_spaced_operator_runs = bool(math_cfg.get("forbid_spaced_operator_runs", False))
    math_font_violations: list[dict[str, Any]] = []
    math_upright_text_font_violations: list[dict[str, Any]] = []
    math_operator_font_violations: list[dict[str, Any]] = []
    spaced_math_operator_runs: list[dict[str, Any]] = []
    math_run_font_counts: dict[str, int] = {}
    if math_cfg.get("require_direct_math_run_font", False):
        for run in body.xpath(".//m:r", namespaces=NS):
            text = math_run_text(run)
            r_fonts = run.find("w:rPr/w:rFonts", namespaces=NS)
            if r_fonts is None:
                item = {"text": text[:80], "ascii": None, "hAnsi": None, "eastAsia": None, "cs": None}
                math_font_violations.append(item)
                warnings.append({"code": "math_run_font_missing", **item})
                continue
            item = {
                "text": text[:80],
                "ascii": r_fonts.get(qn("w:ascii")),
                "hAnsi": r_fonts.get(qn("w:hAnsi")),
                "eastAsia": r_fonts.get(qn("w:eastAsia")),
                "cs": r_fonts.get(qn("w:cs")),
            }
            if item["ascii"]:
                math_run_font_counts[item["ascii"]] = math_run_font_counts.get(item["ascii"], 0) + 1
            actual_fonts = {item[key] for key in ["ascii", "hAnsi", "eastAsia", "cs"]}
            for font in sorted(actual_fonts):
                if font in forbidden_math_fonts:
                    hit = {"location": "math_run", "font": font, "text": text[:80]}
                    forbidden_math_font_hits.append(hit)
                    warnings.append({"code": "forbidden_math_font", **hit})
            if allowed_math_run_fonts and (None in actual_fonts or not actual_fonts.issubset(allowed_math_run_fonts)):
                math_font_violations.append(item)
                warnings.append({"code": "math_run_font_mismatch", "expected": sorted(allowed_math_run_fonts), **item})
            if upright_text_math_font and math_run_prefers_text_font(run) and any(item[key] != upright_text_math_font for key in ["ascii", "hAnsi", "eastAsia", "cs"]):
                math_upright_text_font_violations.append(item)
                warnings.append({"code": "math_upright_text_font_mismatch", "expected": upright_text_math_font, **item})
            if is_operator_math_text(text):
                if operator_math_font and any(item[key] != operator_math_font for key in ["ascii", "hAnsi", "eastAsia", "cs"]):
                    math_operator_font_violations.append(item)
                    warnings.append({"code": "math_operator_font_mismatch", "expected": operator_math_font, **item})
                if forbid_spaced_operator_runs and text != "".join(text.split()):
                    spaced_math_operator_runs.append(item)
                    warnings.append({"code": "spaced_math_operator_run", **item})
    math_count = len(body.xpath(".//m:oMath|.//m:oMathPara", namespaces=NS))
    subscript_count = len(body.xpath(".//m:sSub|.//m:sSup|.//m:sSubSup", namespaces=NS))
    simple_numeric_omml: list[str] = []
    if math_cfg.get("forbid_simple_numeric_omml", False):
        for omml in body.xpath(".//m:oMath", namespaces=NS):
            text = "".join(omml.xpath(".//m:t/text()", namespaces=NS)).strip()
            if text and re.fullmatch(r"[0-9.％% ]+", text):
                simple_numeric_omml.append(text)
                warnings.append({"code": "simple_numeric_omml", "text": text})
    if math_cfg.get("min_omml_count") is not None and math_count < int(math_cfg["min_omml_count"]):
        warnings.append({"code": "low_omml_count", "actual": math_count, "expected_min": math_cfg["min_omml_count"]})
    if math_cfg.get("min_subscript_count") is not None and subscript_count < int(math_cfg["min_subscript_count"]):
        warnings.append({"code": "low_math_subscript_count", "actual": subscript_count, "expected_min": math_cfg["min_subscript_count"]})

    formula_fragment_checks: list[dict[str, Any]] = []
    for spec in math_cfg.get("required_formula_fragments", []):
        name = spec.get("name") or spec.get("anchor") or "formula"
        anchor = compact_text(str(spec.get("anchor", "")))
        matched = False
        check = {"name": name, "anchor": spec.get("anchor", ""), "matched": False, "problems": []}
        for p in body.xpath(".//w:p", namespaces=NS):
            text = paragraph_text(p)
            if anchor and anchor not in compact_text(text):
                continue
            matched = True
            check["matched"] = True
            for token, expected_count in (spec.get("must_contain_counts") or {}).items():
                actual_count = text.count(token)
                if actual_count < int(expected_count):
                    check["problems"].append({"token": token, "actual": actual_count, "expected_min": int(expected_count)})
            for xml_name, min_count in {
                "m:f": spec.get("min_fraction_count"),
                "m:nary": spec.get("min_nary_count"),
                "m:rad": spec.get("min_radical_count"),
                "m:sSub": spec.get("min_subscript_count"),
                "m:sSup": spec.get("min_superscript_count"),
                "m:sSubSup": spec.get("min_subsup_count"),
            }.items():
                if min_count is None:
                    continue
                actual_count = len(p.xpath(f".//{xml_name}", namespaces=NS))
                if actual_count < int(min_count):
                    check["problems"].append({"node": xml_name, "actual": actual_count, "expected_min": int(min_count)})
            break
        if not matched:
            check["problems"].append({"anchor": spec.get("anchor", ""), "actual": 0, "expected_min": 1})
        formula_fragment_checks.append(check)
        if check["problems"]:
            warnings.append({"code": "formula_fragment_incomplete", **check})

    if fix:
        document_tree.write(str(root_dir / "word/document.xml"), encoding="UTF-8", xml_declaration=True, standalone=True)
        styles_tree.write(str(root_dir / "word/styles.xml"), encoding="UTF-8", xml_declaration=True, standalone=True)
        for name, tree in headers.items():
            tree.write(str(root_dir / "word" / name), encoding="UTF-8", xml_declaration=True, standalone=True)

    return {
        "ok": not warnings,
        "warning_count": len(warnings),
        "warnings": warnings,
        "fixes": fixes,
        "metrics": {
            "paragraphs": len(paragraphs),
            "heading_counts": heading_counts,
            "heading_effective_toggles": heading_effective_toggles,
            "heading_toggle_noise": heading_toggle_noise,
            "citation_superscript_violations": citation_superscript_violations,
            "style_gallery_checks": style_gallery_checks,
            "conversion_style_hits": conversion_style_hits,
            "caption_count": caption_count,
            "drawing_count": drawing_count,
            "max_drawing_inches": [round(max_drawing_width, 3), round(max_drawing_height, 3)],
            "header_checks": header_checks,
            "reference_heading_found_at": ref_start,
            "duplicate_reference_numbers": duplicate_reference_numbers,
            "reference_word_wrap_hits": reference_word_wrap_hits,
            "reference_hard_break_hits": reference_hard_break_hits,
            "reference_text_run_count_violations": reference_text_run_count_violations,
            "front_empty_breaks": front_empty_breaks,
            "forbidden_field_code_hits": forbidden_field_code_hits,
            "latin_font_violations": latin_font_violations,
            "configured_math_fonts": configured_math_fonts,
            "forbidden_math_font_hits": forbidden_math_font_hits,
            "math_font_violations": math_font_violations,
            "math_upright_text_font_violations": math_upright_text_font_violations,
            "math_operator_font_violations": math_operator_font_violations,
            "spaced_math_operator_runs": spaced_math_operator_runs,
            "math_run_font_counts": math_run_font_counts,
            "math_count": math_count,
            "math_subscript_count": subscript_count,
            "simple_numeric_omml": simple_numeric_omml,
            "formula_fragment_checks": formula_fragment_checks,
        },
    }


def unpack_docx(docx: Path, root_dir: Path) -> None:
    with zipfile.ZipFile(docx) as zf:
        zf.extractall(root_dir)


def rezip_docx(root_dir: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(root_dir.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(root_dir).as_posix())


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit and optionally repair reusable OOXML thesis invariants.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--config", type=Path)
    parser.add_argument("--json-out", type=Path)
    parser.add_argument("--fix-out", type=Path, help="Write a repaired DOCX copy. The input DOCX is never modified.")
    parser.add_argument("--fail-on-warnings", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    with tempfile.TemporaryDirectory(prefix="ooxml-thesis-guard-") as td:
        root_dir = Path(td) / "docx"
        root_dir.mkdir()
        unpack_docx(args.docx, root_dir)
        result = audit_unpacked(root_dir, cfg, fix=args.fix_out is not None)
        if args.fix_out is not None:
            rezip_docx(root_dir, args.fix_out)

    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    return 1 if args.fail_on_warnings and result["warning_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
