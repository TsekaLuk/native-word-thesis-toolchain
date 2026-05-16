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
6. Run reusable OOXML guardrails before visual inspection. Start from the nearest school config and change only template-specific parameters:
   ```bash
   python scripts/ooxml_thesis_guard.py build/native.docx \
     --config examples/jou-ooxml-guard.json \
     --json-out build/ooxml-guard-report.json \
     --fail-on-warnings
   ```
   For common repairable issues, write a repaired copy rather than editing the input in place:
   ```bash
   python scripts/ooxml_thesis_guard.py build/native.docx \
     --config examples/jou-ooxml-guard.json \
     --fix-out build/native.fixed.docx
   ```
7. Render through a real office engine and inspect page images:
   ```bash
   nwt render-pages build/native.docx build/native-pages.pdf --engine pages
   ```

## Reusable Toolchain

- `scripts/ooxml_thesis_guard.py`: template-aware OOXML audit and optional repair. It checks heading levels/toggle noise, native Chinese style-gallery names and visibility, caption and image centering without inherited indents, header bottom borders/spacers/NBSP page labels, duplicate reference numbering, front-matter empty break paragraphs, exposed field-code artifacts, Latin run fonts, drawing size ceilings, and OMML/subscript/simple-numeric math invariants.
- `examples/jou-ooxml-guard.json`: JOU-style baseline config. Copy it per school or per customer; adjust the header text, spacer length, title/caption rules, figure size caps, and math thresholds instead of rewriting XPath.
- Use `--json-out` as the handoff artifact. It gives stable metrics such as heading counts, caption count, drawing count, max drawing size, reference heading location, and math object counts, so future reviews can compare runs rather than relying on screenshots only.
- Use `--fix-out` for low-risk structural repairs: clearing passive bold/italic toggles, adding direct header bottom borders, zeroing header/caption/image indents, and centering affected paragraphs. Keep school-specific layout construction in the project generator, not in the guard script.

## What To Fix First

- References: remove automatic list numbering before adding `[n]`.
- Captions: center them and use `keepNext` for table captions.
- Tables: center tables, add three-line borders when appropriate, and disable row splitting.
- Long appendix/API tables: do not globally force `cantSplit`; allow row splitting for tables that must span pages, and set fixed semantic column widths.
- Math: use OMML for formulas and narrow symbolic table cells; repair collapsed Pandoc formulas by rebuilding fractions, roots, summation limits, and formula numbers as native OMML.
- Math typography: do not treat font fallback as a fix. Use semantic font roles for generated OMML: keep `settings.xml` `m:mathFont` on a WPS/Word-stable engine font such as `Cambria Math`, use a sharper math font such as `STIX Two Math` only on symbolic variable/operator runs, and force formula-internal upright text (`m:nor`, long metric names such as `Precision`, `Recall`, `NDCG@10`, `Top10`) to `Cambria Math`. Guard both the allowed direct math-run font set and the stricter upright-text font rule; never force `STIX Two Math` onto `m:nor` text inside fractions because WPS can visually drop those words.
- Math formulas: add structural formula guards for high-risk metrics, not only font checks. For F1/Precision/NDCG/CTR-style equations, assert required text fragments, fraction/n-ary/radical node counts, and direct math-run fonts so a formula that still has an equation box but lost numerator/denominator details fails before handoff.
- Math: do not convert plain numeric prose such as `100`, `50`, or `0.08` into standalone OMML objects. Pure-number math boxes show up as strange selectable equation widgets in WPS; flatten them back to ordinary text while keeping real formulas as OMML.
- Math table cells: verify `m:sSub`/`m:sSup` counts, not only `m:oMath` counts. A cell can contain OMML while still rendering as plain italic text if tuple-style symbols such as `("s", "vec")` are serialized as separate math runs instead of one subscript object.
- Math table cells: render-test any cell that mixes prose wrapping and OMML. A line height that is fine for 10.5pt text can clip or hide a trailing formula run such as `λ=0.05`; use formula-safe line spacing for math-bearing cells, and keep tiny explanatory constants in ordinary text when they are part of prose rather than a standalone equation.
- Caption sample sizes: render `N=...` as native math when it appears in table or figure titles.
- Headings: clear inherited italics from Heading 3 and appendix subheadings.
- Headings: for native Word output, do not add extra bold to undergraduate Heading 1/2/3 when the style already uses `黑体`/SimHei. LaTeX templates may use fake bold to compensate for CJK font behavior, but native Word should keep the heading style as black font without `w:b`/`w:bCs` at either style or run level unless the accepted Word template explicitly requires bold.
- Headings: remove passive toggle noise, not only effective formatting. Some Word/WPS pipelines serialize `w:i w:val="0"` or similar false-state nodes on heading runs/styles; even when visually false, they can make toolbar state and later edits look italic. Strip `w:b`/`w:bCs`/`w:i`/`w:iCs` entirely from Heading 1/2/3 styles and heading runs when the target template only needs SimHei.
- Captions: also clear passive bold/italic toggle noise from the caption style after python-docx styling. A leftover `w:bCs` can make Chinese figure/table captions appear or edit as bold even if `style.font.bold = False` was set earlier.
- Captions and generated lists: never let hidden `TC` or `TOC \h \z \f` field markers survive into the delivery copy. WPS can expose those markers as literal text near captions or list titles (`TC "图..." \f F \l 1`). Build visible 附图/附表清单 entries directly or with safe hyperlinks, and run the OOXML guard's field-code checks before handoff.
- Style gallery: do not leave conversion-style English names such as `Normal`, `Heading 1`, `Default Paragraph Font`, or visible unused `Heading 4`. Preserve native Word `styleId`s (`Normal`, `Heading1`, `Heading2`, `Heading3`, `Caption`) for outline/TOC compatibility, but localize `w:name`/aliases to Chinese display names such as `正文`, `标题 1`, `标题 2`, `标题 3`, `题注`, and hide unused Heading 4-9/character styles from the quick style gallery.
- Backmatter headings: match school spacing literally. For this JOU-style target, render `致谢` as `致  谢` and `参考文献` as `参 考 文 献`, while reference-section detection should use compacted text so numbering cleanup still works after display spacing is applied.
- Undergraduate thesis headings: avoid fourth-level structure unless the handbook explicitly requires it; demote LaTeX `\subsubsection`/Word `Heading 4` to bold inline lead-ins and verify no outline level 4 remains.
- Heading numbers: verify visible body heading text, not only the TOC. Some Word/WPS/Pandoc pipelines preserve `Heading 1/2/3` styles but drop visible numbering; rebuild `1`, `1.1`, `1.1.1` from heading levels and skip unnumbered titles such as TOC, conclusion, acknowledgements, references, and appendices.
- Undergraduate thesis TOCs often should list only first- and second-level headings even when the body keeps third-level headings. Diff the accepted senior thesis before using `TOC \o "1-3"`; for JOU-style undergraduate drafts, prefer `TOC \o "1-2"` and visually confirm no `1.1.1` entries appear.
- Main TOC: compare against an accepted same-school DOCX at the XML level. For JOU-style files, the title should be a centered `目    录` paragraph, TOC entries should use native field runs (`HYPERLINK` + nested `PAGEREF`) with dot leaders, visible heading numbers should keep two spaces after the numeric prefix, and the right tab stop may differ from the body/header tab by a few twips (seen: 8302 for main TOC, 8312 for lists/header).
- Lists of tables/figures: do not append them directly under the main TOC. School templates often require `附表清单：` and `附图清单：` to start on separate pages, with title paragraphs having no first-line indent, and entries formatted as `表 3-1 ... [dot leader] page` / `图 3-1 ... [dot leader] page`. Pull page numbers from LaTeX `.aux`/Word fields when possible, and verify the visible rendered pages because a hidden field marker alone is not a usable list.
- Lists of tables/figures: match the accepted thesis typography, not just the text. JOU senior-thesis samples use larger list titles (seen: 14pt bold) and 12pt entries; adding a blank paragraph after `附表清单：` / `附图清单：` can be necessary for the same visual rhythm.
- Front matter: copy or reconstruct from authoritative Word templates; never redraw complex cover pages by eye when handbook assets exist.
- Cover table fidelity is a trade-off between template XML and rendered one-page fit. First try cloning the accepted template table and replacing only run text/underlines. If the cloned table spills to a second page after real student titles are inserted, prefer a stable one-page native table with the correct font (`楷体_GB2312` 四号), bottom-aligned values, and line-adjacent text over a pixel-like XML clone that renders across pages.
- Cover/declaration transitions: inspect manual `w:br w:type="page"` paragraphs after the cover date. If the cover template already fills the page, that extra break can create a blank second page before the declaration.
- Front-matter blank pages: normalize inherited `w:pageBreakBefore` in declaration/authorization paragraphs. Keep only deliberate page starts; duplicate a section break plus `pageBreakBefore` before the abstract can render as a blank page in WPS even when Pages tolerates it. A standalone empty paragraph containing only `w:sectPr` before the abstract can also become a visible blank page in Word/WPS; move that section break onto the previous real paragraph.
- Front-matter page starts: when a copied template already carries section breaks or page-sized frames, prefer a single inline `w:br w:type="page"` at the first real paragraph over `w:pageBreakBefore`; Word/WPS can otherwise materialize an unexpected blank page between declaration/authorization and abstracts.
- Abstract pages: if the handbook uses bordered abstract frames, rebuild Chinese/English abstracts as template-style table frames rather than plain paragraphs.
- Abstract frames: the thesis title inside the bordered Chinese/English abstract frame is part of the template typography; make it bold as well, not only the outer page title.
- Abstract frames: undergraduate Chinese/English abstracts should usually be one continuous body paragraph. Collapse source paragraph breaks inside the abstract body during Word generation, while keeping `关键词：` / `Keywords:` as separate lines. In the English abstract frame, force ASCII, hAnsi, and eastAsia font bindings to `Times New Roman` for English runs.
- Abstract frames: copy exact `w:trHeight` and `w:hRule` from the handbook/reference for Chinese and English abstract frames; generic `atLeast` heights often look close in XML but produce shorter bordered boxes in WPS/Word.
- Abstract page typography: keep the outer Chinese abstract title in `黑体`; the `关键词：` / `Keywords:` line should not have first-line indentation even when body abstract paragraphs do.
- Abstract title sizing must be render-tested with the bordered frame. Blindly matching a larger XML `w:sz` can push the English abstract frame to the next page in Pages/WPS. Prefer the largest title size that keeps title and frame on the same rendered page; no blank/title-only page is acceptable.
- Cover date and signature lines: inspect inherited `w:ind`, `w:tabs`, and `w:framePr`, not just `w:jc`; a paragraph can be "centered" inside a shifted frame or first-line indent.
- Cover underline fields: values sitting on horizontal lines should keep the handbook font size (for JOU, `楷体_GB2312` 四号 = 14pt / `w:sz=28`) and use bottom-aligned table cells with near-zero bottom margins; centered vertical alignment makes text float above the underline.
- Headers: compare against the handbook/reference rendering, not only page fields. JOU body headers should use gray text, left-aligned thesis label, right-aligned `第 X 页 共 Y 页` via a right tab stop, and a gray bottom border line under the header paragraph.
- Headers: put the gray bottom border directly on the generated header paragraph as well as in the header style. WPS/Word template inheritance can drop or hide style-only borders when sections are regenerated; XML validation should confirm `header*.xml` has `w:pPr/w:pBdr/w:bottom`.
- Headers: prefer native right-tab layout over literal spaces between the school label and `第 X 页 共 Y 页`: set a paragraph `w:tabs/w:tab w:val="right"` at the body content width, insert one `w:tab`, and keep the page label pieces joined with NBSP. Literal spaces are only a last-resort template clone; they can wrap when both current and total page numbers become two digits.
- Headers: explicitly zero header paragraph indentation (`left=0`, `right=0`, `firstLine=0`) even when the referenced header style appears to have no indentation. WPS can show inherited first-line indent in header-edit mode unless the header paragraph overrides it directly.
- Headers: test both one-digit and two-digit page numbers. A spacer that fits `第 1 页 共 50 页` can wrap at `第 10 页 共 50 页`; reserve two-digit width and use non-breaking spaces inside the right-side page label so WPS cannot break between `50` and `页`.
- Header parameters: when matching an accepted JOU senior thesis, diff both `word/header*.xml` and `word/styles.xml`. A visually similar direct `Header` paragraph can still be wrong if the reference uses a numeric style such as `styleId=14` with its own bottom border, tab stops, snap-to-grid, and run fonts. Match the reference style, paragraph `jc`, run font, and field-run formatting instead of approximating with a new header style.
- Header total-page fields are engine-sensitive. `SECTIONPAGES` can render as `1` in Pages even when `PAGE` starts correctly. If WPS/Pages cannot calculate the body-section total reliably, use a verified static body-page total in the header result and document that it must be refreshed after pagination-changing edits.
- Page setup: diff `w:pgMar`/`w:pgSz`/header distance against an accepted same-school `.docx`, not only centimeter guesses. For JOU reference files seen here, body pages used A4 `11906 x 16838` with margins `top=1440 right=1797 bottom=1440 left=1797`, so TOC/list/header right tab stops should align to the resulting content width, not an arbitrary 8500 twip.
- Field safety: do not enable global `w:updateFields` until the target Office engine is verified. It can trigger Word security prompts or make Pages rebuild TOCs and break pagination.
- Screenshot grids and multi-image tables: if a figure is implemented as a borderless table of images, put the figure caption in a final row of the same borderless table. A separate following caption paragraph can split across pages because Word cannot `keepNext` a table to a following paragraph reliably.
- Figure and caption alignment: centered figures/captions must explicitly clear paragraph indentation (`left=0`, `right=0`, `firstLine=0`, no `hanging`) before setting `jc=center`. Do not rely on a centered style alone; WPS/Word can center inside an inherited first-line-indent frame, making the image or caption look offset.
- Oversized figures/tables: do not allow a single figure/table to occupy an otherwise blank page unless the handbook explicitly permits a plate page. For LaTeX, cap tall images around `height=0.68\textheight` and set a high `floatpagefraction` (for example `0.92`) to discourage float-only pages. For native Word, cap oversized drawing extents after conversion and verify with rendered pages that the caption and surrounding text remain on the same page.

## When More Detail Is Needed

- Read `references/checklist.md` for acceptance checks and common failure signatures.
- Read `references/config.md` when adapting YAML math table repair to a new thesis.

## Delivery Rule

Do not call the Word deliverable complete until both are true:

- `nwt validate` has no blocking failures.
- A rendered PDF from Pages, Word, WPS, or LibreOffice has been inspected on the affected pages.
