from __future__ import annotations

import os
from pathlib import Path
from typing import Any


SKIP_DIRS = {
    ".git",
    ".hg",
    ".svn",
    ".venv",
    "__pycache__",
    ".pytest_cache",
    ".archive",
    "node_modules",
    "artifacts",
    "bin",
    "dist",
    "build",
    "obj",
    "tmp",
    "out",
    "coverage",
}

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".svg", ".pdf"}
DOC_SUFFIXES = {".doc", ".docx", ".pdf"}
TEMPLATE_KEYWORDS = ("模板", "template", "样式", "格式", "封面", "声明", "授权")
HANDBOOK_KEYWORDS = ("手册", "规范", "要求", "说明", "guide", "manual", "毕业论文")
REFERENCE_KEYWORDS = ("定稿", "参考", "学长", "学姐", "reference", "final")
NON_THESIS_DELIVERABLE_KEYWORDS = ("开题", "答辩", "外文翻译", "presentation", "slides")


def project_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [name for name in dirnames if name not in SKIP_DIRS]
        base = Path(dirpath)
        for filename in filenames:
            if filename == ".DS_Store":
                continue
            files.append(base / filename)
    return sorted(files)


def rel(path: Path, root: Path) -> str:
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return path.as_posix()


def has_keyword(path: Path, keywords: tuple[str, ...]) -> bool:
    haystack = path.as_posix().casefold()
    return any(keyword.casefold() in haystack for keyword in keywords)


def has_project_keyword(path: Path, root: Path, keywords: tuple[str, ...]) -> bool:
    return has_keyword(Path(rel(path, root)), keywords)


def is_thesis_deliverable_candidate(path: Path, root: Path) -> bool:
    return not has_project_keyword(path, root, NON_THESIS_DELIVERABLE_KEYWORDS)


def preferred_tex(files: list[Path]) -> list[Path]:
    tex_files = [path for path in files if path.suffix.lower() == ".tex"]
    preferred_names = {"main.tex", "thesis.tex", "paper.tex"}
    return sorted(tex_files, key=lambda path: (path.name.casefold() not in preferred_names, len(path.parts), path.as_posix()))


def collect(files: list[Path], root: Path) -> dict[str, list[str]]:
    latex = preferred_tex(files)
    docx_files = [path for path in files if path.suffix.lower() == ".docx"]
    pdf_files = [
        path
        for path in files
        if path.suffix.lower() == ".pdf" and (path.name.casefold() == "main.pdf" or is_thesis_deliverable_candidate(path, root))
    ]
    format_docs = [path for path in docx_files if is_thesis_deliverable_candidate(path, root)]
    categories = {
        "latex_main": latex[:3],
        "bibliography": [path for path in files if path.suffix.lower() == ".bib"],
        "compiled_pdf": pdf_files,
        "draft_docx": [
            path
            for path in format_docs
            if not has_project_keyword(path, root, TEMPLATE_KEYWORDS + REFERENCE_KEYWORDS + HANDBOOK_KEYWORDS)
        ],
        "template_docx": [path for path in format_docs if has_project_keyword(path, root, TEMPLATE_KEYWORDS)],
        "handbook": [
            path
            for path in files
            if path.suffix.lower() in DOC_SUFFIXES
            and is_thesis_deliverable_candidate(path, root)
            and has_project_keyword(path, root, HANDBOOK_KEYWORDS)
        ],
        "reference_docx": [path for path in format_docs if has_project_keyword(path, root, REFERENCE_KEYWORDS)],
        "figure_assets": [
            path
            for path in files
            if path.suffix.lower() in IMAGE_SUFFIXES
            and any(part.casefold() in {"figure", "figures", "image", "images", "assets", "插图", "图片"} for part in path.parts)
        ],
    }
    return {key: [rel(path, root) for path in value[:20]] for key, value in categories.items()}


def audit_project(root: Path) -> dict[str, Any]:
    root = root.resolve()
    files = project_files(root)
    found = collect(files, root)
    source_ready = bool(found["latex_main"] or found["draft_docx"] or found["compiled_pdf"])
    governance_ready = bool(found["template_docx"] or found["handbook"] or found["reference_docx"])
    warnings: list[str] = []
    if not found["latex_main"]:
        warnings.append("未发现 main/thesis/paper.tex；若只有 PDF 或 DOCX，需要先确认是否可回到 LaTeX/结构化源文件。")
    if not found["template_docx"]:
        warnings.append("未发现疑似学校模板 DOCX；封面、声明、授权书和样式库需要外部模板输入。")
    if not found["handbook"]:
        warnings.append("未发现疑似毕业论文手册/规范；页边距、页眉、目录、题注等参数缺少权威来源。")
    if not found["reference_docx"]:
        warnings.append("未发现疑似学长/定稿参考 DOCX；像素级对齐会缺少渲染参照。")
    if not found["figure_assets"]:
        warnings.append("未发现集中图片/插图资产目录；image2/Py 图表字体治理可能需要先归档图片来源。")

    commands = ["nwt intake . --json build/intake-report.json"]
    if found["latex_main"]:
        tex = found["latex_main"][0]
        parent = Path(tex).parent.as_posix()
        resource_path = "." if parent == "." else parent
        commands.append(f"nwt draft-latex {tex} build/draft.docx --resource-path {resource_path}")
    elif found["draft_docx"]:
        commands.append(f"nwt polish {found['draft_docx'][0]} build/native.docx --config examples/jou-thesis.yaml")
    else:
        commands.append("先补齐 LaTeX 源文件或可编辑 DOCX；不要把 PDF 转 Word 当作正式交付主路径。")
    commands.extend(
        [
            "nwt polish build/draft.docx build/native.docx --config examples/jou-thesis.yaml",
            "nwt validate build/native.docx --config examples/jou-thesis.yaml --json build/native-report.json",
            "python scripts/ooxml_thesis_guard.py build/native.docx --config examples/jou-ooxml-guard.json --json-out build/ooxml-guard-report.json --fail-on-warnings",
            "nwt render-pages build/native.docx build/native-pages.pdf --engine pages",
        ]
    )
    return {
        "root": root.as_posix(),
        "file_count": len(files),
        "ok": source_ready and governance_ready,
        "source_ready": source_ready,
        "governance_ready": governance_ready,
        "found": found,
        "warnings": warnings,
        "recommended_commands": commands,
    }
