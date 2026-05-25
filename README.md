# 📄 PDF-Word-to-Markdown

**PDF/Word ↔ Markdown 双向转换工具** · 公式 → LaTeX · 表格 · 图片 · 双栏布局

Bidirectional PDF & DOCX to Markdown converter with LaTeX formula, table, image & dual-column preservation.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-green.svg)]()

---

[English](#english) | [中文](#chinese)

---

## <a name="english"></a>English

### Quick Start

```bash
# 1. Clone
git clone https://github.com/xangfeng/pdf-word-to-markdown.git
cd pdf-word-to-markdown

# 2. Install dependencies
pip install -r requirements.txt

# 3. Convert PDF to Markdown
python -m pdf_word_to_markdown.pdf_to_md paper.pdf -o output/

# 4. Convert Word to Markdown
python -m pdf_word_to_markdown.word_to_md document.docx

# 5. Reconstruct Word from Markdown
python -m pdf_word_to_markdown.md_to_word output.md output_format.json
```

### Features

#### PDF → Markdown

| Feature | Implementation | Engine |
|---------|---------------|--------|
| Layout Detection | PP-DocLayoutV2 (24 block types) | RapidDoc |
| Text OCR | RapidOCR (Chinese/English, ONNX) | RapidDoc |
| Formula Recognition | PP-FormulaNet_plus → `$$...$$` / `$...$` | RapidDoc |
| Table Recognition | SLANET/UNITABLE → HTML → Markdown | RapidDoc |
| Reading Order | Coordinate clustering + column-aware algorithm | Self-developed |
| Column Layout | Auto-detect + `::: columns` markers | Self-developed |
| Formula Enhancement | Vision model: image → LaTeX (optional) | Kimi/OpenAI/DeepSeek |

#### Word → Markdown

- Paragraph styles → heading levels (auto-infer 1/1.1/1.1.1 numbering)
- OMML formulas → LaTeX `$...$` / `$$...$$`
- Tables → Markdown format
- Images → extracted to `images/`
- Format metadata → `format.json`

#### Markdown → Word Reconstruction

- Format JSON drives style rebuilding (font/size/alignment/spacing/indent)
- LaTeX formulas → OMML (native Word math equations)
- Table restoration (including merged cells)
- Image insertion
- Dual-column section breaks

### Performance

| Document | Pages | First Run | Cached |
|----------|-------|-----------|--------|
| Dual-column academic PDF | 12 | ~80s | **~38s** |

First run downloads ONNX models (~100MB, one-time). Subsequent runs use cached models. Pure CPU inference (ONNX Runtime), no GPU required.

### Installation

**Requirements:** Python 3.10+

```bash
pip install -r requirements.txt
```

Dependencies:
- `rapid-doc[cpu]` — PDF parsing engine (ONNX)
- `python-docx` — Word document I/O
- `PyMuPDF` — PDF rendering
- `mistune` — Markdown AST parsing
- `Pillow`, `opencv-python`, `numpy` — Image processing
- `latex2mathml` — LaTeX ↔ MathML conversion
- `openai` — API calls (optional, for vision enhancement)

### Optional: Vision Formula Enhancement (Multi-Provider)

Default: OFF. To enable, set an API key in `.env`:

```bash
# Copy config template
cp .env.example .env

# Edit .env, fill in ONE of these:
MOONSHOT_API_KEY=sk-xxx     # Kimi (recommended: moonshot-v1-8k-vision-preview, cheap)
OPENAI_API_KEY=sk-xxx       # OpenAI (gpt-4o-mini)
DEEPSEEK_API_KEY=sk-xxx     # DeepSeek (vision permission required)
ANTHROPIC_API_KEY=sk-xxx    # Claude (claude-haiku-4-5)
```

Then run with:

```bash
# Auto-detect provider
python -m pdf_word_to_markdown.pdf_to_md paper.pdf --formulas

# Or specify provider
python -m pdf_word_to_markdown.pdf_to_md paper.pdf --formulas --provider kimi
```

### Output Format

Each conversion produces:
- `{name}.md` — Markdown with `$...$`/`$$...$$` formulas, tables, image refs, `::: columns` markers
- `{name}_format.json` — Format metadata (styles, fonts, layout info)
- `images/` — Extracted images

### Architecture / 架构

```
pdf_word_to_markdown/
├── pdf_to_md.py               # PDF → MD pipeline / PDF转Markdown调度层
├── word_to_md.py              # Word → MD (python-docx + OMML) / Word转Markdown
├── md_to_word.py              # MD → Word reconstruction / Markdown还原Word
├── ooml_latex.py              # OMML ↔ LaTeX bidirectional / 公式双向转换
├── utils.py                   # Shared utilities / 共享工具（公式清理、标题推断）
└── engines/
    ├── rapid_doc_engine.py     # Primary PDF engine (ONNX/CPU) / 主引擎
    ├── layout_reader_engine.py # Column-aware reading order / 分栏阅读顺序
    ├── docling_engine.py       # Table enhancement (optional) / 表格增强
    └── formula_vision_engine.py # Vision formula (multi-provider) / 视觉公式
```

### File Descriptions / 文件说明

| File 文件 | Description 说明 |
|-----------|-----------------|
| `pdf_to_md.py` | PDF → Markdown orchestrator. Calls RapidDoc engine for layout/OCR/formula/table, then layout_reader for column ordering. Supports `--formulas` (vision API) and `--tables` (Docling) flags. |
| `word_to_md.py` | Word (.docx) → Markdown. Uses python-docx to read paragraphs/tables/images, extracts OMML formulas as LaTeX, generates format.json with full style metadata. |
| `md_to_word.py` | Markdown + format.json → Word (.docx). Reconstructs document with original styles, converts LaTeX formulas back to OMML via MathML, handles tables/images/columns. |
| `ooml_latex.py` | OMML ↔ LaTeX bidirectional conversion. Parses Office Math XML to LaTeX (OMML→LaTeX) and converts LaTeX→MathML→OMML for Word insertion. |
| `utils.py` | Shared helpers: LaTeX space cleaning (`clean_latex`), heading level inference from numbering patterns (`infer_heading_level`), color conversion, image extraction, format JSON I/O. |
| `engines/rapid_doc_engine.py` | Primary PDF engine. Wraps RapidDoc's ONNX pipeline: PP-DocLayoutV2 layout detection, PP-FormulaNet_plus formula recognition, SLANET/UNITABLE table recognition, RapidOCR text. Extracts blocks with bbox/type/text, converts HTML tables to Markdown. |
| `engines/layout_reader_engine.py` | Column-aware reading order. X-coordinate clustering to detect columns, assigns each block to a column, sorts top-to-bottom within columns and left-to-right across columns. Handles full-width (span) blocks like figures. Renders `::: columns` Markdown. |
| `engines/docling_engine.py` | Docling (IBM) table enhancement. Uses TableFormer for best-in-class table structure recovery (cell-adjacency F1=0.89). Optional, enabled with `--tables`. |
| `engines/formula_vision_engine.py` | Multi-provider vision API for formula enhancement. Supports Kimi/OpenAI/DeepSeek/Anthropic/Custom via unified interface. Auto-detects configured provider from environment variables. Default: OFF. |

---

## <a name="chinese"></a>中文

### 快速开始

```bash
# 1. 克隆仓库
git clone https://github.com/xangfeng/pdf-word-to-markdown.git
cd pdf-word-to-markdown

# 2. 安装依赖
pip install -r requirements.txt

# 3. PDF 转 Markdown
python -m pdf_word_to_markdown.pdf_to_md paper.pdf -o output/

# 4. Word 转 Markdown
python -m pdf_word_to_markdown.word_to_md document.docx

# 5. Markdown 还原为 Word
python -m pdf_word_to_markdown.md_to_word output.md output_format.json
```

### 功能特性

#### PDF → Markdown

| 功能 | 实现方式 | 引擎 |
|------|---------|------|
| 版面检测 | PP-DocLayoutV2（24 种区块类型） | RapidDoc |
| 文字识别 | RapidOCR（中英文，ONNX） | RapidDoc |
| 公式识别 | PP-FormulaNet_plus → `$$...$$` / `$...$` | RapidDoc |
| 表格识别 | SLANET/UNITABLE → HTML → Markdown | RapidDoc |
| 阅读顺序 | 坐标聚类 + 栏感知算法 | 自研 |
| 分栏布局 | 自动检测 + `::: columns` 标记 | 自研 |
| 公式增强 | 视觉模型：公式图片 → LaTeX（可选） | Kimi/OpenAI/DeepSeek |

#### Word → Markdown

- 段落样式 → 标题层级（自动推断 1/1.1/1.1.1 编号模式）
- OMML 公式 → LaTeX `$...$` / `$$...$$`
- 表格 → Markdown 格式
- 图片 → 提取到 `images/` 目录
- 格式信息 → `format.json`

#### Markdown → Word 还原

- 格式 JSON 驱动样式重建（字体/字号/对齐/行距/缩进）
- LaTeX 公式 → OMML（Word 原生数学公式）
- 表格还原（含合并单元格）
- 图片插入
- 双栏分节

### 性能参考

| 文档类型 | 页数 | 首次运行 | 缓存后 |
|---------|------|---------|--------|
| 双栏学术论文 PDF | 12 页 | ~80s | **~38s** |

首次运行需下载 ONNX 模型（约 100MB，仅一次），后续使用缓存。纯 CPU 推理（ONNX Runtime），无需 GPU。

### 环境要求

**Python 3.10+**

```bash
pip install -r requirements.txt
```

核心依赖说明：
- `rapid-doc[cpu]` — PDF 解析引擎（ONNX 推理）
- `python-docx` — Word 文档读写
- `PyMuPDF` — PDF 渲染与解析
- `mistune` — Markdown AST 解析
- `Pillow`、`opencv-python`、`numpy` — 图像处理
- `latex2mathml` — LaTeX ↔ MathML 公式转换
- `openai` — API 调用（可选，视觉公式增强用）

### 可选：多模态视觉公式增强（支持多厂商）

默认关闭。如需启用，在 `.env` 文件中配置 API key：

```bash
# 复制配置模板
cp .env.example .env

# 编辑 .env，任选一个填入：
MOONSHOT_API_KEY=sk-xxx     # Kimi（推荐：moonshot-v1-8k-vision-preview，省 token）
OPENAI_API_KEY=sk-xxx       # OpenAI（gpt-4o-mini）
DEEPSEEK_API_KEY=sk-xxx     # DeepSeek（需开通视觉权限）
ANTHROPIC_API_KEY=sk-xxx    # Claude（claude-haiku-4-5）
```

然后运行：

```bash
# 自动检测提供商
python -m pdf_word_to_markdown.pdf_to_md paper.pdf --formulas

# 或指定提供商
python -m pdf_word_to_markdown.pdf_to_md paper.pdf --formulas --provider kimi
```

### 输出格式

每次转换生成：
- `{name}.md` — Markdown 正文（含 `$...$`/`$$...$$` 公式、表格、图片引用、`::: columns` 分栏标记）
- `{name}_format.json` — 格式文档（样式、字体、布局信息）
- `images/` — 提取的图片

### 项目架构

```
pdf_word_to_markdown/
├── pdf_to_md.py               # PDF → MD 调度层
├── word_to_md.py              # Word → MD（python-docx + OMML）
├── md_to_word.py              # MD → Word 还原
├── ooml_latex.py              # OMML ↔ LaTeX 双向转换
├── utils.py                   # 共享工具（公式清理/标题推断）
└── engines/
    ├── rapid_doc_engine.py     # PDF 主引擎（ONNX/CPU）
    ├── layout_reader_engine.py # 分栏阅读顺序引擎
    ├── docling_engine.py       # Docling 表格增强（可选）
    └── formula_vision_engine.py # 视觉公式增强（多厂商/默认关）
```

### 常见问题

**Q: 首次运行为什么这么慢？**
A: 需要从 ModelScope 下载 ONNX 模型文件（约 100MB），仅第一次。后续运行直接使用缓存。

**Q: 需要 GPU 吗？**
A: 不需要。全部使用 ONNX Runtime CPU 推理。

**Q: 公式识别不准怎么办？**
A: 启用视觉增强（`--formulas`），用 Kimi/OpenAI 等视觉模型将公式图片进一步识别为 LaTeX。

**Q: 支持扫描版 PDF 吗？**
A: RapidOCR 内置文字检测，对清晰扫描件支持良好。严重模糊/扭曲的扫描件效果会下降。

---

## Acknowledgments / 致谢

本项目基于以下开源项目构建。详见 [ACKNOWLEDGMENTS.md](ACKNOWLEDGMENTS.md)。

- [RapidDoc](https://github.com/RapidAI/RapidDoc) — Apache 2.0 · ONNX document parsing pipeline
- [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR) — Apache 2.0 · OCR engine
- [Docling](https://github.com/docling-project/docling) — MIT · IBM document parsing
- [LayoutReader](https://github.com/ppaanngggg/layoutreader) — Reading order model
- [Marker](https://github.com/VikParuchuri/marker) — GPL · PDF→MD reference
- [MinerU](https://github.com/opendatalab/MinerU) — Apache 2.0 · Document parsing reference
- [MarkItDown](https://github.com/microsoft/markitdown) — MIT · Microsoft file→MD converter
- [MarkEverything](https://pypi.org/project/mark-everything/) — MIT · OMML formula reference

## License

MIT License. See [LICENSE](LICENSE).

Third-party libraries retain their original licenses (Apache 2.0, MIT, GPL, etc.).

---

<p align="center">
  <b>🚧 本项目仍在持续优化中，如有不足欢迎提出 Issues 或 PR！</b><br>
  <i>This project is under active development. Feedback and contributions are welcome!</i>
</p>
