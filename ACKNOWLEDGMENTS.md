# 致谢 / Acknowledgments

本项目（pdf_word_to_markdown）基于以下优秀开源项目构建。我们向其作者和维护者致以诚挚的感谢。

## 直接集成

### [RapidDoc](https://github.com/RapidAI/RapidDoc)
- **许可**: Apache License 2.0
- **用途**: PDF 文档解析主引擎。PP-DocLayoutV2 版面检测、PP-FormulaNet_plus 公式识别、SLANET/UNITABLE 表格识别、RapidOCR 文字识别。
- **致谢**: RapidAI 团队提供了高性能的 ONNX 文档解析管线。

### [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR)
- **许可**: Apache License 2.0
- **用途**: PP-OCRv4/v5 文字识别模型，中文 OCR 的核心基础。
- **致谢**: 百度 PaddlePaddle 团队。

### [Docling](https://github.com/docling-project/docling)
- **许可**: MIT License
- **用途**: DocLayNet 版面分析、TableFormer 表格结构还原（可选增强）。
- **致谢**: IBM Research 团队。

## 算法参考

### [LayoutReader](https://github.com/ppaanngggg/layoutreader)
- **许可**: 开源
- **用途**: 阅读顺序预测模型，为分栏阅读顺序算法提供启发。
- **致谢**: hantian 及 FreeOCR-AI 团队。

### [pdf_oxide](https://github.com/yfedoseev/pdf_oxide)
- **许可**: 开源
- **用途**: ColumnAware 分栏检测算法思路参考（Issue #270）。
- **致谢**: yfedoseev。

## 功能参考

### [Marker](https://github.com/VikParuchuri/marker)
- **许可**: GPL License
- **用途**: PDF → Markdown 转换流程参考。Surya OCR、Texify 公式识别、Tabled 表格识别。
- **致谢**: VikParuchuri。

### [MinerU](https://github.com/opendatalab/MinerU)
- **许可**: Apache License 2.0
- **用途**: 文档解析流程架构参考。doclayout_yolo 版面检测、UniMERNet 公式识别。
- **致谢**: OpenDataLab（上海人工智能实验室）。

### [MarkItDown](https://github.com/microsoft/markitdown)
- **许可**: MIT License
- **用途**: Word → Markdown 转换方案参考。
- **致谢**: Microsoft。

### [MarkEverything](https://github.com/your-org/mark-everything)
- **许可**: MIT License
- **用途**: OMML 公式 → LaTeX 转换方案参考。
- **致谢**: mark-everything 团队。

## 核心依赖

本项目的功能实现离不开以下 Python 库：

- [python-docx](https://github.com/python-openxml/python-docx) — MIT License · Word 文档读写
- [PyMuPDF](https://github.com/pymupdf/PyMuPDF) — AGPL License · PDF 渲染与解析
- [mistune](https://github.com/lepture/mistune) — BSD License · Markdown AST 解析
- [latex2mathml](https://github.com/roniemartinez/latex2mathml) — MIT License · LaTeX → MathML 转换
- [numpy](https://github.com/numpy/numpy) — BSD License · 数值计算
- [OpenAI Python SDK](https://github.com/openai/openai-python) — Apache 2.0 License · API 调用

## 视觉模型

本项目可选集成的视觉公式增强功能支持以下厂商的 API：

- [Kimi (Moonshot AI)](https://platform.kimi.com) · moonshot-v1-8k-vision-preview / kimi-k2.6
- [OpenAI](https://platform.openai.com) · GPT-4o / GPT-4o-mini
- [DeepSeek](https://platform.deepseek.com) · DeepSeek-V4-Pro / V4-Flash
- [Anthropic](https://console.anthropic.com) · Claude Haiku / Sonnet / Opus

## 声明

本项目的 MIT 许可证仅适用于我们自行编写的代码。集成的第三方库和模型各自保留其原始许可证。请在使用前查阅各项目的许可证条款。
