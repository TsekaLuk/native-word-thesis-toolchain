# Native Word Thesis Checklist

## Blocking Checks

- DOCX zip integrity passes.
- Cover, declaration, authorization, abstract, TOC, body, references, and appendix all exist.
- Body section has page header and page-number fields where required.
- Front-matter sections should not accidentally inherit the正文 header; check section-level `headerReference` and `pgNumType` rather than only rendered body pages.
- References render as `[1] ...`, not `1. [1] ...`.
- Figure captions are below figures; table captions are above tables.
- Table captions are centered and keep with the following table.
- Tables are centered and rows do not split across pages unless explicitly allowed.
- Equations are OMML/native Word math where editability is expected.
- Rendered formulas should preserve native structures such as fractions, roots, summation limits, and formula numbers; Pandoc-generated math that visually collapses into a plain text stream must be rewritten with OMML builders.
- Appendix symbol cells do not show vertical fragments like `s`, `_`, `vec`.
- Sample-size fragments in captions such as `N=250 samples` should use OMML for `N=250` when the school expects formula-grade typography.
- Heading 3 and appendix subheadings are bold regular text, not italic.
- No raw test paths, stack traces, class paths, TODOs, or unresolved placeholders remain.

## Common Failure Signatures

- `reference_numPr_left > 0`: Word still has automatic list numbering in bibliography.
- `heading_italic_left` non-empty: converted direct `w:i` formatting survived.
- `floating_table_captions` non-empty: table caption can detach from table.
- Low `math_object_count` after a formula-heavy paper: formulas may have become plain text or images.
- Rendered PDF differs from DOCX checks: trust the rendered PDF and fix layout, because office engines differ.

## Visual Pages To Inspect

- First cover/front-matter pages.
- Chinese and English abstracts.
- TOC and figure/table lists.
- First body chapter page.
- A representative figure page.
- A dense table page.
- A formula page.
- First reference page.
- Appendix pages with API/path tables or math tables.
