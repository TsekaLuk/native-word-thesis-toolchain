from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Any

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W, "r": R, "m": M}


def qname(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def w_el(tag: str, text: str | None = None, **attrs: str) -> etree._Element:
    el = etree.Element(qname(W, tag))
    for key, value in attrs.items():
        el.set(qname(W if key.startswith("w_") else R, key.split("_", 1)[1]), value)
    if text is not None:
        el.text = text
    return el


def m_el(tag: str, text: str | None = None) -> etree._Element:
    el = etree.Element(qname(M, tag))
    if text is not None:
        el.text = text
    return el


def node_text(node: etree._Element) -> str:
    return "".join(node.xpath(".//w:t/text() | .//m:t/text()", namespaces=NS)).strip()


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def ensure_ppr(p: etree._Element) -> etree._Element:
    p_pr = p.find("w:pPr", namespaces=NS)
    if p_pr is None:
        p_pr = etree.Element(qname(W, "pPr"))
        p.insert(0, p_pr)
    return p_pr


def ensure_tblpr(tbl: etree._Element) -> etree._Element:
    tbl_pr = tbl.find("w:tblPr", namespaces=NS)
    if tbl_pr is None:
        tbl_pr = etree.Element(qname(W, "tblPr"))
        tbl.insert(0, tbl_pr)
    return tbl_pr


def set_jc(container: etree._Element, value: str) -> None:
    p_pr = ensure_ppr(container) if container.tag == qname(W, "p") else ensure_tblpr(container)
    jc = p_pr.find("w:jc", namespaces=NS)
    if jc is None:
        jc = etree.SubElement(p_pr, qname(W, "jc"))
    jc.set(qname(W, "val"), value)


def set_keep(p: etree._Element, *, keep_next: bool = False, keep_lines: bool = True) -> None:
    p_pr = ensure_ppr(p)
    if keep_next and p_pr.find("w:keepNext", namespaces=NS) is None:
        p_pr.insert(0, etree.Element(qname(W, "keepNext")))
    if keep_lines and p_pr.find("w:keepLines", namespaces=NS) is None:
        p_pr.insert(0, etree.Element(qname(W, "keepLines")))


def set_spacing(p: etree._Element, *, before: str = "0", after: str = "0", line: str = "300") -> None:
    p_pr = ensure_ppr(p)
    spacing = p_pr.find("w:spacing", namespaces=NS)
    if spacing is None:
        spacing = etree.SubElement(p_pr, qname(W, "spacing"))
    spacing.set(qname(W, "before"), before)
    spacing.set(qname(W, "after"), after)
    spacing.set(qname(W, "line"), line)
    spacing.set(qname(W, "lineRule"), "auto")


def remove_paragraph_numbering(p: etree._Element) -> None:
    p_pr = ensure_ppr(p)
    for num_pr in p_pr.xpath("./w:numPr", namespaces=NS):
        p_pr.remove(num_pr)
    p_style = p_pr.find("w:pStyle", namespaces=NS)
    if p_style is not None:
        style = (p_style.get(qname(W, "val")) or "").lower()
        if style.startswith(("list", "bibliography")) or style in {"a", "a1"}:
            p_pr.remove(p_style)


def set_reference_paragraph(p: etree._Element) -> None:
    remove_paragraph_numbering(p)
    p_pr = ensure_ppr(p)
    set_jc(p, "left")
    ind = p_pr.find("w:ind", namespaces=NS)
    if ind is None:
        ind = etree.SubElement(p_pr, qname(W, "ind"))
    ind.attrib.clear()
    ind.set(qname(W, "left"), "720")
    ind.set(qname(W, "hanging"), "420")
    set_spacing(p)


def replace_paragraph_text(p: etree._Element, text: str) -> None:
    runs = p.xpath("./w:r", namespaces=NS)
    if not runs:
        p.append(make_text_run(text))
        return
    written = False
    for run in runs:
        texts = run.xpath("./w:t", namespaces=NS)
        if texts and not written:
            texts[0].text = text
            for extra in texts[1:]:
                extra.text = ""
            written = True
        else:
            for t in texts:
                t.text = ""


def rewrite_paragraph_inline_math(p: etree._Element, parts: list[tuple[str, str | list[Any]]]) -> None:
    p_pr = p.find("w:pPr", namespaces=NS)
    preserved = [deepcopy(p_pr)] if p_pr is not None else []
    for child in list(p):
        p.remove(child)
    for child in preserved:
        p.append(child)
    r_pr = make_text_rpr()
    for kind, value in parts:
        if kind == "text":
            if value:
                p.append(make_text_run(str(value), r_pr))
        elif kind == "math":
            p.append(make_omath(value if isinstance(value, list) else [str(value)]))


def replace_sample_size_with_math(p: etree._Element) -> bool:
    text = node_text(p)
    matches = list(re.finditer(r"([（(])N\s*=\s*(\d+)\s*samples([）)])", text))
    if not matches:
        return False
    parts: list[tuple[str, str | list[Any]]] = []
    pos = 0
    for match in matches:
        parts.append(("text", text[pos : match.start()]))
        parts.append(("text", match.group(1)))
        parts.append(("math", ["N=", match.group(2)]))
        parts.append(("text", " samples" + match.group(3)))
        pos = match.end()
    parts.append(("text", text[pos:]))
    rewrite_paragraph_inline_math(p, parts)
    return True


def make_text_rpr(
    east_asia: str = "宋体",
    latin: str = "Times New Roman",
    size_half_points: str = "21",
    bold: bool = False,
) -> etree._Element:
    r_pr = etree.Element(qname(W, "rPr"))
    fonts = etree.SubElement(r_pr, qname(W, "rFonts"))
    fonts.set(qname(W, "ascii"), latin)
    fonts.set(qname(W, "hAnsi"), latin)
    fonts.set(qname(W, "eastAsia"), east_asia)
    if bold:
        etree.SubElement(r_pr, qname(W, "b"))
    etree.SubElement(r_pr, qname(W, "color")).set(qname(W, "val"), "000000")
    etree.SubElement(r_pr, qname(W, "sz")).set(qname(W, "val"), size_half_points)
    etree.SubElement(r_pr, qname(W, "szCs")).set(qname(W, "val"), size_half_points)
    return r_pr


def make_text_run(text: str, r_pr: etree._Element | None = None) -> etree._Element:
    r = etree.Element(qname(W, "r"))
    if r_pr is not None:
        r.append(deepcopy(r_pr))
    t = etree.SubElement(r, qname(W, "t"))
    t.text = text
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def make_math_run(text: str) -> etree._Element:
    r = m_el("r")
    t = m_el("t", text)
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def make_math_text(text: str) -> etree._Element:
    r = m_el("r")
    r_pr = m_el("rPr")
    r_pr.append(m_el("nor"))
    r.append(r_pr)
    t = m_el("t", text)
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def make_math_container(tag: str, parts: list[Any]) -> etree._Element:
    node = m_el(tag)
    append_math_parts(node, parts)
    return node


def make_math_subscript(base: str, subscript: str, *, normal_base: bool = False) -> etree._Element:
    node = m_el("sSub")
    node.append(m_el("sSubPr"))
    base_node = m_el("e")
    base_node.append(make_math_text(base) if normal_base else make_math_run(base))
    sub_node = m_el("sub")
    sub_node.append(make_math_run(subscript))
    node.append(base_node)
    node.append(sub_node)
    return node


def make_math_subsup(base_parts: list[Any], sub_parts: list[Any], sup_parts: list[Any]) -> etree._Element:
    node = m_el("sSubSup")
    node.append(m_el("sSubSupPr"))
    node.append(make_math_container("e", base_parts))
    node.append(make_math_container("sub", sub_parts))
    node.append(make_math_container("sup", sup_parts))
    return node


def make_math_fraction(num_parts: list[Any], den_parts: list[Any]) -> etree._Element:
    node = m_el("f")
    node.append(m_el("fPr"))
    node.append(make_math_container("num", num_parts))
    node.append(make_math_container("den", den_parts))
    return node


def make_math_radical(parts: list[Any]) -> etree._Element:
    node = m_el("rad")
    rad_pr = m_el("radPr")
    deg_hide = m_el("degHide")
    deg_hide.set(qname(M, "val"), "1")
    rad_pr.append(deg_hide)
    node.append(rad_pr)
    node.append(m_el("deg"))
    node.append(make_math_container("e", parts))
    return node


def make_math_nary_sum(sub_parts: list[Any], sup_parts: list[Any], body_parts: list[Any]) -> etree._Element:
    node = m_el("nary")
    nary_pr = m_el("naryPr")
    chr_node = m_el("chr")
    chr_node.set(qname(M, "val"), "∑")
    nary_pr.append(chr_node)
    lim_loc = m_el("limLoc")
    lim_loc.set(qname(M, "val"), "undOvr")
    nary_pr.append(lim_loc)
    if not sub_parts:
        sub_hide = m_el("subHide")
        sub_hide.set(qname(M, "val"), "1")
        nary_pr.append(sub_hide)
    if not sup_parts:
        sup_hide = m_el("supHide")
        sup_hide.set(qname(M, "val"), "1")
        nary_pr.append(sup_hide)
    node.append(nary_pr)
    node.append(make_math_container("sub", sub_parts))
    node.append(make_math_container("sup", sup_parts))
    node.append(make_math_container("e", body_parts))
    return node


def math_parts(value: Any) -> list[Any]:
    return value if isinstance(value, list) else [value]


def append_math_parts(parent: etree._Element, parts: list[Any]) -> None:
    for part in parts:
        if isinstance(part, etree._Element):
            parent.append(part)
        elif isinstance(part, dict) and "text" in part:
            parent.append(make_math_text(str(part["text"])))
        elif isinstance(part, dict) and "sub" in part:
            base, sub = part["sub"]
            parent.append(make_math_subscript(str(base), str(sub), normal_base=bool(part.get("normal_base"))))
        elif isinstance(part, dict) and "subsup" in part:
            base, sub, sup = part["subsup"]
            parent.append(make_math_subsup(math_parts(base), math_parts(sub), math_parts(sup)))
        elif isinstance(part, dict) and "frac" in part:
            num, den = part["frac"]
            parent.append(make_math_fraction(math_parts(num), math_parts(den)))
        elif isinstance(part, dict) and "rad" in part:
            parent.append(make_math_radical(math_parts(part["rad"])))
        elif isinstance(part, dict) and "sum" in part:
            spec = part["sum"]
            if isinstance(spec, dict):
                sub = math_parts(spec.get("sub", []))
                sup = math_parts(spec.get("sup", []))
                body = math_parts(spec.get("body", []))
            else:
                sub, sup, body = spec
                sub = math_parts(sub)
                sup = math_parts(sup)
                body = math_parts(body)
            parent.append(make_math_nary_sum(sub, sup, body))
        else:
            parent.append(make_math_run(str(part)))


def make_omath(parts: list[Any]) -> etree._Element:
    math = m_el("oMath")
    append_math_parts(math, parts)
    return math


def parse_caption(text: str) -> tuple[str, str, str] | None:
    match = re.match(r"^([图表])((?:\d+(?:\.\d+|-\s*\d+)|[A-Z]\.\d+))\s+(.+)$", text)
    if not match:
        return None
    kind, number, caption = match.groups()
    return kind, number.replace("- ", ".").replace("-", "."), caption


def format_caption(p: etree._Element, *, keep_next: bool) -> None:
    remove_paragraph_numbering(p)
    set_jc(p, "center")
    set_keep(p, keep_next=keep_next, keep_lines=True)
    p_pr = ensure_ppr(p)
    ind = p_pr.find("w:ind", namespaces=NS)
    if ind is None:
        ind = etree.SubElement(p_pr, qname(W, "ind"))
    ind.set(qname(W, "firstLine"), "0")
    set_spacing(p, before="60", after="60")


def is_real_table_caption(children: list[etree._Element], idx: int) -> bool:
    for node in children[idx + 1 :]:
        if node.tag == qname(W, "tbl"):
            return True
        if node.tag == qname(W, "p") and node_text(node):
            return False
    return False


def is_real_figure_caption(children: list[etree._Element], idx: int) -> bool:
    for node in reversed(children[:idx]):
        if node.tag == qname(W, "p") and node.xpath(".//w:drawing|.//w:pict", namespaces=NS):
            return True
        if node.tag in {qname(W, "tbl"), qname(W, "p")} and node_text(node):
            return False
    return False


def center_table_and_lock_rows(tbl: etree._Element) -> None:
    set_jc(tbl, "center")
    for tr in tbl.xpath("./w:tr", namespaces=NS):
        tr_pr = tr.find("w:trPr", namespaces=NS)
        if tr_pr is None:
            tr_pr = etree.Element(qname(W, "trPr"))
            tr.insert(0, tr_pr)
        if tr_pr.find("w:cantSplit", namespaces=NS) is None:
            tr_pr.append(etree.Element(qname(W, "cantSplit")))


def clear_heading_italics(body: etree._Element) -> None:
    heading_pattern = re.compile(r"^(?:\d+(?:\.\d+){2,3}|[A-Z]\.\d+)\s+")
    for p in body.xpath("./w:p", namespaces=NS):
        text = node_text(p)
        style_vals = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
        style = style_vals[0] if style_vals else ""
        if style in {"Heading3", "3", "30"} or heading_pattern.match(text):
            for r_pr in p.xpath(".//w:rPr", namespaces=NS):
                for tag in ("i", "iCs"):
                    for node in r_pr.xpath(f"./w:{tag}", namespaces=NS):
                        r_pr.remove(node)


def normalize_references(body: etree._Element, config: dict[str, Any]) -> None:
    heading = config.get("references", {}).get("heading", "参考文献")
    stop_prefixes = tuple(config.get("references", {}).get("stop_prefixes", ["附录", "Appendix"]))
    in_refs = False
    number = 1
    for p in body.xpath("./w:p", namespaces=NS):
        text = node_text(p)
        if text == heading:
            in_refs = True
            continue
        if any(text.startswith(prefix) for prefix in stop_prefixes):
            in_refs = False
        if not in_refs or not text:
            continue
        cleaned = re.sub(r"\s+", " ", text).strip()
        if not re.match(r"^[\[\［]\d+[\]\］]", cleaned):
            cleaned = f"[{number}] {cleaned}"
        replace_paragraph_text(p, cleaned)
        set_reference_paragraph(p)
        number += 1


def set_cell_borders(tc_pr: etree._Element, top=None, bottom=None, left=None, right=None) -> None:
    borders = tc_pr.find("w:tcBorders", namespaces=NS)
    if borders is None:
        borders = etree.SubElement(tc_pr, qname(W, "tcBorders"))
    for side, spec in {"top": top, "bottom": bottom, "left": left, "right": right}.items():
        node = borders.find(f"w:{side}", namespaces=NS)
        if node is None:
            node = etree.SubElement(borders, qname(W, side))
        if spec is None:
            node.set(qname(W, "val"), "nil")
        else:
            color, size = spec
            node.set(qname(W, "val"), "single")
            node.set(qname(W, "sz"), str(size))
            node.set(qname(W, "space"), "0")
            node.set(qname(W, "color"), color)


def make_table_cell(value: Any, width: int, *, header: bool, align: str = "center", no_wrap: bool = False) -> etree._Element:
    tc = etree.Element(qname(W, "tc"))
    tc_pr = etree.SubElement(tc, qname(W, "tcPr"))
    tc_w = etree.SubElement(tc_pr, qname(W, "tcW"))
    tc_w.set(qname(W, "w"), str(width))
    tc_w.set(qname(W, "type"), "dxa")
    etree.SubElement(tc_pr, qname(W, "vAlign")).set(qname(W, "val"), "center")
    if no_wrap:
        etree.SubElement(tc_pr, qname(W, "noWrap"))
    p = etree.SubElement(tc, qname(W, "p"))
    p_pr = etree.SubElement(p, qname(W, "pPr"))
    etree.SubElement(p_pr, qname(W, "jc")).set(qname(W, "val"), align)
    spacing = etree.SubElement(p_pr, qname(W, "spacing"))
    spacing.set(qname(W, "before"), "0")
    spacing.set(qname(W, "after"), "0")
    spacing.set(qname(W, "line"), "300")
    spacing.set(qname(W, "lineRule"), "auto")
    r_pr = make_text_rpr(bold=header)
    append_cell_value(p, value, r_pr)
    return tc


def append_cell_value(p: etree._Element, value: Any, r_pr: etree._Element) -> None:
    if isinstance(value, dict) and "math" in value:
        p.append(make_omath(value["math"]))
        return
    if isinstance(value, list):
        for part in value:
            if isinstance(part, dict) and "math" in part:
                p.append(make_omath(part["math"]))
            else:
                p.append(make_text_run(str(part), r_pr))
        return
    p.append(make_text_run(str(value), r_pr))


def make_configured_table(spec: dict[str, Any]) -> etree._Element:
    widths = [int(width) for width in spec["widths"]]
    rows = spec["rows"]
    tbl = etree.Element(qname(W, "tbl"))
    tbl_pr = etree.SubElement(tbl, qname(W, "tblPr"))
    tbl_w = etree.SubElement(tbl_pr, qname(W, "tblW"))
    tbl_w.set(qname(W, "w"), str(sum(widths)))
    tbl_w.set(qname(W, "type"), "dxa")
    etree.SubElement(tbl_pr, qname(W, "jc")).set(qname(W, "val"), "center")
    etree.SubElement(tbl_pr, qname(W, "tblLayout")).set(qname(W, "type"), "fixed")
    grid = etree.SubElement(tbl, qname(W, "tblGrid"))
    for width in widths:
        etree.SubElement(grid, qname(W, "gridCol")).set(qname(W, "w"), str(width))
    for row_idx, row in enumerate(rows):
        tr = etree.SubElement(tbl, qname(W, "tr"))
        tr_pr = etree.SubElement(tr, qname(W, "trPr"))
        etree.SubElement(tr_pr, qname(W, "cantSplit"))
        for col_idx, value in enumerate(row):
            cell = make_table_cell(
                value,
                widths[col_idx],
                header=row_idx == 0,
                align="center" if col_idx in set(spec.get("center_columns", [])) else "left",
                no_wrap=col_idx in set(spec.get("nowrap_columns", [])),
            )
            borders = {}
            if row_idx == 0:
                borders["top"] = ("000000", 12)
                borders["bottom"] = ("000000", 6)
            if row_idx == len(rows) - 1:
                borders["bottom"] = ("000000", 12)
            set_cell_borders(cell.find("w:tcPr", namespaces=NS), **borders)
            tr.append(cell)
    return tbl


def replace_table_after_caption(body: etree._Element, caption: str, new_table: etree._Element) -> bool:
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag != qname(W, "p") or node_text(child) != caption:
            continue
        for candidate in children[idx + 1 :]:
            if candidate.tag == qname(W, "tbl"):
                pos = body.index(candidate)
                body.remove(candidate)
                body.insert(pos, new_table)
                return True
            if candidate.tag == qname(W, "p") and node_text(candidate):
                return False
    return False


def replace_configured_math_tables(body: etree._Element, config: dict[str, Any]) -> list[str]:
    replaced: list[str] = []
    for spec in config.get("math_tables", []):
        caption = spec["caption"]
        if replace_table_after_caption(body, caption, make_configured_table(spec)):
            replaced.append(caption)
    return replaced


def repair_caption_sample_math(body: etree._Element) -> int:
    count = 0
    for p in body.xpath("./w:p", namespaces=NS):
        if parse_caption(node_text(p)) and replace_sample_size_with_math(p):
            count += 1
    return count


def polish_document_xml(root: etree._Element, config: dict[str, Any]) -> dict[str, Any]:
    body = root.xpath("//w:body", namespaces=NS)[0]
    normalize_references(body, config)
    replaced_tables = replace_configured_math_tables(body, config)
    sample_caption_math_count = 0
    if config.get("native_sample_size_captions", True):
        sample_caption_math_count = repair_caption_sample_math(body)
    clear_heading_italics(body)
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag == qname(W, "tbl"):
            center_table_and_lock_rows(child)
            if idx > 0 and children[idx - 1].tag == qname(W, "p"):
                parsed = parse_caption(node_text(children[idx - 1]))
                if parsed and parsed[0] == "表":
                    format_caption(children[idx - 1], keep_next=True)
            continue
        if child.tag != qname(W, "p"):
            continue
        parsed = parse_caption(node_text(child))
        if not parsed:
            continue
        kind = parsed[0]
        if kind == "表" and is_real_table_caption(children, idx):
            format_caption(child, keep_next=True)
        elif kind == "图" and is_real_figure_caption(children, idx):
            format_caption(child, keep_next=False)
    return {"replaced_math_tables": replaced_tables, "sample_caption_math_count": sample_caption_math_count}


def enable_update_fields(unpacked: Path) -> None:
    settings_path = unpacked / "word/settings.xml"
    if not settings_path.exists():
        return
    tree = etree.parse(str(settings_path))
    root = tree.getroot()
    if not root.xpath("./w:updateFields", namespaces=NS):
        node = etree.Element(qname(W, "updateFields"))
        node.set(qname(W, "val"), "true")
        root.insert(0, node)
    tree.write(str(settings_path), encoding="UTF-8", xml_declaration=True, standalone=True)


def rezip_dir(src_dir: Path, out_path: Path) -> None:
    if out_path.exists():
        out_path.unlink()
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(src_dir.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(src_dir).as_posix())


def polish_docx(input_path: Path, output_path: Path, config: dict[str, Any]) -> dict[str, Any]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="native-word-thesis-") as td:
        unpacked = Path(td) / "docx"
        unpacked.mkdir()
        with zipfile.ZipFile(input_path) as zf:
            zf.extractall(unpacked)
        document_path = unpacked / "word/document.xml"
        tree = etree.parse(str(document_path))
        report = polish_document_xml(tree.getroot(), config)
        tree.write(str(document_path), encoding="UTF-8", xml_declaration=True, standalone=True)
        enable_update_fields(unpacked)
        rezip_dir(unpacked, output_path)
    return report


def copy_skill_to(skill_source: Path, destination_root: Path) -> Path:
    dest = destination_root / skill_source.name
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(skill_source, dest)
    return dest
