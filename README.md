# Native Word Thesis Toolchain

![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/github/license/TsekaLuk/native-word-thesis-toolchain)
![Workflow](https://img.shields.io/badge/workflow-LaTeX%20%2B%20OOXML%20%2B%20render%20check-2f6fed)

![Native Word Thesis Toolchain cover](assets/readme-hero.png)

> Stop shipping "converted Word files". Ship native, inspectable thesis documents.

把 LaTeX/PDF 论文交付件转成老师能直接检查的 Word 原生论文。目标不是“能打开的 docx”，而是封面、样式、题注、公式、目录、页眉、参考文献都能被 Word/WPS 原生编辑和审查的交付件。

这个仓库沉淀的是通用工具链，不保存客户论文、学校模板、参考论文或私有图片。模板、手册、参考定稿都作为本地输入传入。

## 为什么要用它

大多数论文转 Word 的失败，不是卡在“转换”这一步，而是卡在老师会挑的细节：

- PDF 转 Word 后，公式像图片或碎文本，无法原生编辑。
- LaTeX/Pandoc 生成的 `s_vec`、`s_final` 在窄表格里竖排、拆字。
- 图注、表注看起来居中，实际带着首行缩进或没有跟随图表。
- 参考文献变成 `1. [1]` 双序号，正文上标引用缺少方括号。
- 页眉、页码、目录、附图清单、附表清单在 Word/WPS/Pages 里渲染不一致。
- 封面、声明、授权书靠肉眼复刻，日期、横线、签名区总有一两个像素不对。

`native-word-thesis-toolchain` 把这些反复踩过的坑变成可复用命令、OOXML guard 和回归检查。

## 适合谁

- 需要把 LaTeX 本科论文交成 Word 初稿的人。
- 已经有 PDF，但老师明确说 PDF 不行、必须交 Word 原生版本的人。
- 正在维护学校论文模板、毕业论文代交付流程、论文排版自动化工具的人。
- 想把一次次手工修版沉淀成可重复工具链的人。

## 60 秒跑第一轮审计

```bash
python3 -m pip install -e .
nwt intake /path/to/thesis-project --json build/intake-report.json
```

`intake` 会先判断新论文是否具备结构化源文件、学校模板/手册、学长定稿、插图资产和参考文献。没有这些输入时，它会明确阻止“凭感觉修 Word”的低质量路径。

## 标准流水线

1. 从 LaTeX 生成 Word 草稿：

```bash
nwt draft-latex thesis/main.tex build/draft.docx --resource-path thesis
```

2. 按学校配置修复为 Word 原生排版：

```bash
nwt polish build/draft.docx build/native.docx --config examples/jou-thesis.yaml
```

3. 做结构检查：

```bash
nwt validate build/native.docx --config examples/jou-thesis.yaml --json build/native-report.json
```

4. 跑 OOXML 细节护栏：

```bash
python scripts/ooxml_thesis_guard.py build/native.docx \
  --config examples/jou-ooxml-guard.json \
  --json-out build/ooxml-guard-report.json \
  --fail-on-warnings
```

5. 在 macOS 上用真实 Office 引擎导出 PDF 视觉回归：

```bash
nwt render-pages build/native.docx build/native-pages.pdf --engine pages
```

## 它会守住什么

- **公式原生性**：生成 OMML，检查分式、根号、求和、上下标、公式字体和纯数字伪公式。
- **题注与图表**：图注/表注居中，清除首行缩进，表题 `keepNext`，表格居中并避免错误拆行。
- **参考文献**：移除自动编号，避免双序号，保持条目连续，检查正文上标引用的 `[ ]`。
- **页眉页码**：检查 section 继承、页眉灰色下边框、右侧页码宽度、单双位页码换行风险。
- **样式库**：保留 Word 原生样式语义，同时把显示名、Heading 4、Pandoc 残留样式纳入 guard。
- **前置页**：封面、声明、授权书、摘要页不靠截图兜底，优先从权威模板或手册迁移结构。
- **渲染验证**：结构检查之后还要用 Pages/Word/WPS/LibreOffice 的真实渲染结果验收。

## 仓库边界

本仓库只放通用工具：

- `src/native_word_thesis/`：CLI、OOXML 修复、渲染、验证、接入审计。
- `scripts/ooxml_thesis_guard.py`：可配置的 OOXML 审计与低风险修复。
- `examples/`：学校级配置样例。
- `skills/native-word-thesis/`：Codex skill，方便下一个论文项目复用。
- `tests/`：把踩过的坑固化成回归测试。

不放客户论文、不放私有模板、不放学校手册原件、不放参考定稿。它们应该作为每个项目的本地输入。

## Skill 安装

仓库内置 Codex skill：

```bash
nwt install-skill
```

安装后，下次可以直接让 Codex 使用 `native-word-thesis` skill 处理新的 LaTeX/PDF/DOCX 论文项目。

## 设计原则

- PDF 转 Word 只作为诊断 fallback；正式交付优先走 LaTeX/Pandoc 草稿 + OOXML 后处理。
- 能结构化检查的细节，不靠截图和主观描述。
- 公式必须是 OMML 或 Word 可编辑对象，不能把应编辑公式做成截图。
- 模板资源不内置，按项目从学校手册、参考定稿、用户提供文件读取。
- 封面、声明、授权书这类前置页不能只信样式编号；日期、签名、封面字段表等关键锚点要做显式结构检查。
- 每轮交付必须同时有结构检查和渲染检查，不能只说“文件能打开”。
