from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


def render_with_pages(input_docx: Path, output_pdf: Path, delay: int = 4) -> None:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    script = f'''
set inputPath to "{input_docx.resolve()}"
set outputPath to "{output_pdf.resolve()}"
set inputFile to POSIX file inputPath
set outputFile to POSIX file outputPath
tell application "Pages"
    activate
    open inputFile
    delay {delay}
    set theDoc to front document
    export theDoc to outputFile as PDF
    close theDoc saving no
end tell
'''
    subprocess.run(["osascript"], input=script, text=True, check=True)


def render_with_libreoffice(input_docx: Path, output_pdf: Path) -> None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice/soffice not found")
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_pdf.parent), str(input_docx)],
        check=True,
    )
    generated = output_pdf.parent / f"{input_docx.stem}.pdf"
    if generated != output_pdf and generated.exists():
        generated.replace(output_pdf)


def render_docx(input_docx: Path, output_pdf: Path, engine: str = "pages") -> None:
    if engine == "pages":
        render_with_pages(input_docx, output_pdf)
    elif engine == "libreoffice":
        render_with_libreoffice(input_docx, output_pdf)
    else:
        raise ValueError(f"Unknown render engine: {engine}")
