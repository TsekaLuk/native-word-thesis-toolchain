# Config Notes

The YAML config is intentionally small. It should describe project-specific repair targets, not the whole thesis.

## References

```yaml
references:
  heading: "参考文献"
  stop_prefixes: ["附录", "Appendix"]
```

The tool will scan paragraphs after the heading, remove Word list numbering, and ensure a single bracket prefix.

## Caption Sample Sizes

```yaml
native_sample_size_captions: true
```

When enabled, table/figure captions such as `表6.3 ...（N=250 samples）` keep the caption text but render `N=250` as OMML math. This avoids reviewer-visible plain-text math in captions and figure/table lists.

## Math Tables

Use `math_tables` to replace a fragile converted table after a matching caption.

```yaml
math_tables:
  - caption: "表A.1 推荐评分各分量权重与计算约定"
    widths: [2100, 1450, 1300, 3900]
    center_columns: [0, 1, 2]
    nowrap_columns: [1, 2]
    rows:
      - ["分量", "符号", "权重", "计算方式"]
      - ["向量语义相似度", {math: [{sub: ["s", "vec"]}]}, "0.35", "余弦相似度"]
```

Supported cell values:

- Plain string: normal Word text.
- `{math: [...]}`: one OMML math object.
- Mixed list: normal text and math objects inside one cell.
- `{sub: ["s", "vec"]}` inside `math`: renders an OMML subscript.

Prefer `≥` and `≤` over `>=` and `<=`; some office engines visually split ASCII comparisons.
