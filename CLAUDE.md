# 文档转换器 Skill

Word/PDF ↔ Markdown 双向转换，保留公式（LaTeX）、表格、图片和双栏布局。

## 触发方式

当用户说"转换这个 PDF/Word 为 Markdown"、"把这个 Markdown 还原为 Word"等时使用。

## 转换命令

### PDF → Markdown

```bash
python -m pdf_word_to_markdown.pdf_to_md <输入.pdf> -o <输出目录>
```

可选参数：
- `--formulas` 启用视觉模型公式增强（需配置 API key）
- `--tables` 启用 Docling 表格增强
- `--lang ch` 中文文档（默认）

### Word → Markdown

```bash
python -m pdf_word_to_markdown.word_to_md <输入.docx> [输出目录]
```

### Markdown → Word 还原

```bash
python -m pdf_word_to_markdown.md_to_word <输入.md> <格式.json> [输出.docx]
```

## 输出格式

每次转换产生：
- `{name}.md` — Markdown 正文（公式 `$...$`/`$$...$$`，表格，图片引用，`::: columns` 分栏标记）
- `{name}_format.json` — 格式文档（样式、字体、布局信息）
- `images/` — 提取的图片

## 首次使用

```bash
pip install -r requirements.txt
```

首次运行会下载 ONNX 模型（约 100MB），后续缓存复用。

## 视觉公式增强（可选）

如需将 PDF 中的公式图片转为 LaTeX，在 `.env` 中填入任一视觉模型 API key：

```bash
MOONSHOT_API_KEY=sk-xxx     # Kimi (推荐)
OPENAI_API_KEY=sk-xxx       # OpenAI
DEEPSEEK_API_KEY=sk-xxx     # DeepSeek
```

然后添加 `--formulas` 参数启用。

## 注意事项

1. 公式默认由 RapidDoc 的 PP-FormulaNet 识别为 LaTeX，视觉增强仅处理漏网之鱼
2. 双栏学术论文自动检测分栏，输出 `::: columns` 标记
3. 格式 JSON 记录了完整样式，用于 Word 还原
4. 性能参考：12 页双栏学术论文，首次 ~80s（下载 ONNX 模型），缓存后 ~40s
