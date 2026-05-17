# Native Word Thesis Toolchain

把 LaTeX/PDF 论文交付件转成“老师能直接检查的 Word 原生论文”的复用工具链。

这个仓库沉淀的是通用工具，不保存客户论文、学校模板、参考论文或私有图片。模板、手册、参考定稿都作为本地输入传入。

## 能解决的常见坑

- PDF 转 Word 后公式、表格、图片漂浮，不能作为原生 Word 初稿交付。
- LaTeX/Pandoc 转出来的 `s_vec`、`s_final` 在窄表格里竖排、拆字。
- 表题里的 `N=250 samples` 这类样本量符号只是普通文本。
- 参考文献变成 `1. [1]` 双序号。
- 图注/表注不居中，或没有跟随图片/表格。
- Word section 继承导致摘要/目录/正文页眉不一致，或页码总数突然变化。
- 封面/声明页从手册或定稿复制后，样式编号语义漂移导致日期、签名行等关键文本错位。
- 表格不是三线表，行被分页拆开。
- Heading 3/附录小节继承了斜体。
- 目录、附图/附表清单、页眉页码、字段更新没有结构化检查。

## 安装

```bash
python3 -m pip install -e .
```

## 基本流程

0. 新论文接入前先做材料审计，确认有没有结构化源文件、学校模板/手册、学长定稿、插图资产和参考文献：

```bash
nwt intake /path/to/thesis-project --json build/intake-report.json
```

这个报告用于决定能否直接进入 LaTeX/Pandoc + OOXML 后处理路径；如果只有 PDF，或缺少模板/手册/定稿参照，应先补齐输入，不要直接开始“凭感觉修 Word”。

1. 从 LaTeX 生成 Word 草稿：

```bash
nwt draft-latex thesis/main.tex build/draft.docx --resource-path thesis
```

2. 按配置修复为 Word 原生排版：

```bash
nwt polish build/draft.docx build/native.docx --config examples/jou-thesis.yaml
```

3. 做结构检查：

```bash
nwt validate build/native.docx --json build/native-report.json
```

4. 在 macOS 上用 Pages 导出 PDF 进行视觉回归：

```bash
nwt render-pages build/native.docx build/native-pages.pdf --engine pages
```

## Skill 安装

仓库内置 Codex skill：

```bash
nwt install-skill
```

安装后，下次可以直接让 Codex 使用 `native-word-thesis` skill 处理新的 LaTeX/PDF 论文项目。

## 设计原则

- PDF 转 Word 只作为诊断 fallback；正式交付优先走 LaTeX/Pandoc 草稿 + OOXML 后处理。
- 公式必须是 OMML 或 Word 可编辑对象，不能把应编辑公式做成截图。
- 模板资源不内置，按项目从学校手册、参考定稿、用户提供文件读取。
- 封面、声明、授权书这类前置页不能只信样式编号；日期、签名、封面字段表等关键锚点要做显式结构检查。
- 每轮交付必须同时有结构检查和渲染检查，不能只说“文件能打开”。
