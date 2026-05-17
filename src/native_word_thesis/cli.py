from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import yaml

from .intake import audit_project
from .ooxml import copy_skill_to, polish_docx
from .render import render_docx
from .validate import dump_report, validate_docx


def load_config(path: Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def cmd_draft_latex(args: argparse.Namespace) -> int:
    pandoc = shutil.which("pandoc")
    if not pandoc:
        print("pandoc not found. Install pandoc first for LaTeX draft conversion.", file=sys.stderr)
        return 2
    cmd = [
        pandoc,
        str(args.input),
        "-o",
        str(args.output),
        "--from",
        "latex",
        "--to",
        "docx",
    ]
    if args.resource_path:
        cmd.extend(["--resource-path", args.resource_path])
    if args.reference_doc:
        cmd.extend(["--reference-doc", str(args.reference_doc)])
    if args.bibliography:
        cmd.extend(["--bibliography", str(args.bibliography)])
    args.output.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(cmd, check=True)
    return 0


def cmd_intake(args: argparse.Namespace) -> int:
    report = audit_project(args.project)
    print(dump_report(report, args.json))
    return 0 if report["ok"] or args.soft else 1


def cmd_polish(args: argparse.Namespace) -> int:
    config = load_config(args.config)
    report = polish_docx(args.input, args.output, config)
    print(dump_report(report))
    return 0


def cmd_validate(args: argparse.Namespace) -> int:
    config = load_config(args.config)
    report = validate_docx(args.input, config)
    print(dump_report(report, args.json))
    return 0 if report["ok"] or args.soft else 1


def cmd_render_pages(args: argparse.Namespace) -> int:
    render_docx(args.input, args.output, args.engine)
    print(args.output)
    return 0


def cmd_install_skill(args: argparse.Namespace) -> int:
    repo_root = Path(__file__).resolve().parents[2]
    source = repo_root / "skills/native-word-thesis"
    codex_home = Path(args.codex_home or os.environ.get("CODEX_HOME", Path.home() / ".codex"))
    dest = copy_skill_to(source, codex_home / "skills")
    print(dest)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="native-word-thesis")
    sub = parser.add_subparsers(dest="command", required=True)

    intake = sub.add_parser("intake", help="Audit whether a new thesis project has the inputs needed for native Word conversion")
    intake.add_argument("project", type=Path)
    intake.add_argument("--json", type=Path)
    intake.add_argument("--soft", action="store_true", help="Always exit 0 while still printing the report")
    intake.set_defaults(func=cmd_intake)

    draft = sub.add_parser("draft-latex", help="Convert a LaTeX source file to a draft .docx with pandoc")
    draft.add_argument("input", type=Path)
    draft.add_argument("output", type=Path)
    draft.add_argument("--resource-path")
    draft.add_argument("--reference-doc", type=Path)
    draft.add_argument("--bibliography", type=Path)
    draft.set_defaults(func=cmd_draft_latex)

    polish = sub.add_parser("polish", help="Post-process a draft .docx into a native Word thesis .docx")
    polish.add_argument("input", type=Path)
    polish.add_argument("output", type=Path)
    polish.add_argument("--config", type=Path)
    polish.set_defaults(func=cmd_polish)

    validate = sub.add_parser("validate", help="Run structural checks on a generated .docx")
    validate.add_argument("input", type=Path)
    validate.add_argument("--config", type=Path)
    validate.add_argument("--json", type=Path)
    validate.add_argument("--soft", action="store_true", help="Always exit 0 while still printing the report")
    validate.set_defaults(func=cmd_validate)

    render = sub.add_parser("render-pages", help="Export a .docx to PDF through a local office engine")
    render.add_argument("input", type=Path)
    render.add_argument("output", type=Path)
    render.add_argument("--engine", choices=["pages", "libreoffice"], default="pages")
    render.set_defaults(func=cmd_render_pages)

    install = sub.add_parser("install-skill", help="Install bundled Codex skill into CODEX_HOME/skills")
    install.add_argument("--codex-home")
    install.set_defaults(func=cmd_install_skill)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
