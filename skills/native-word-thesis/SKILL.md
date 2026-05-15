---
name: native-word-thesis
description: Use when converting or repairing a thesis from LaTeX, PDF, or draft DOCX into a native editable Word thesis with school-template formatting, captions, references, OMML formulas, three-line tables, TOC/list checks, and rendered PDF validation.
---

# Native Word Thesis

Use this skill for thesis Word delivery, especially when a PDF is not acceptable and the initial DOCX has conversion artifacts.

## Default Workflow

1. Locate source artifacts: LaTeX/PDF, generated draft DOCX, school handbook/template, accepted reference thesis, figures, bibliography, and any generated reports.
2. Prefer LaTeX/Pandoc draft generation over PDF-to-Word:
   ```bash
   nwt draft-latex thesis/main.tex build/draft.docx --resource-path thesis
   ```
3. Create or adapt a YAML config for the school/project. Start from `examples/jou-thesis.yaml` in this repo.
4. Polish the draft:
   ```bash
   nwt polish build/draft.docx build/native.docx --config examples/jou-thesis.yaml
   ```
5. Validate structure:
   ```bash
   nwt validate build/native.docx --config examples/jou-thesis.yaml --json build/native-report.json
   ```
6. Render through a real office engine and inspect page images:
   ```bash
   nwt render-pages build/native.docx build/native-pages.pdf --engine pages
   ```

## What To Fix First

- References: remove automatic list numbering before adding `[n]`.
- Captions: center them and use `keepNext` for table captions.
- Tables: center tables, add three-line borders when appropriate, and disable row splitting.
- Math: use OMML for formulas and narrow symbolic table cells.
- Headings: clear inherited italics from Heading 3 and appendix subheadings.
- Front matter: copy or reconstruct from authoritative Word templates; never redraw complex cover pages by eye when handbook assets exist.

## When More Detail Is Needed

- Read `references/checklist.md` for acceptance checks and common failure signatures.
- Read `references/config.md` when adapting YAML math table repair to a new thesis.

## Delivery Rule

Do not call the Word deliverable complete until both are true:

- `nwt validate` has no blocking failures.
- A rendered PDF from Pages, Word, WPS, or LibreOffice has been inspected on the affected pages.
