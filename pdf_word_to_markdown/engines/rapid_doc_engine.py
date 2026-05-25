"""RapidDoc 引擎封装 —— ONNX 纯 CPU 管线 (Apache 2.0)

核心能力：
  PP-DocLayoutV2: 版面检测 (24种block类型) + 阅读顺序
  PP-FormulaNet_plus: 公式识别 → LaTeX ($$...$$)
  SLANET/UNITABLE: 表格识别 → HTML
  RapidOCR: 文字识别 (中文优化)
"""

import json
from pathlib import Path
from html.parser import HTMLParser

from ..utils import clean_latex, infer_heading_level


class RapidDocEngine:
    """RapidDoc 引擎 —— 快速、准确、CPU 友好"""

    def __init__(self, lang: str = "ch"):
        self.lang = lang
        self._model = None

    def _get_model(self):
        if self._model is None:
            from rapid_doc import RapidDoc
            self._model = RapidDoc(
                lang=self.lang,
                formula_enable=True,
                table_enable=True,
                make_md_mode="mm_markdown",
                image_output_mode="url",
                image_dir_name="images",
            )
        return self._model

    def convert(self, pdf_path: str, output_dir: str) -> dict:
        """转换 PDF，返回结构化结果"""
        model = self._get_model()
        model(
            pdf_path,
            output_dir=output_dir,
            f_dump_middle_json=True,
            f_dump_content_list=True,
        )

        output_path = Path(output_dir)
        md_files = list(output_path.glob("**/*.md"))
        json_files = list(output_path.glob("**/*_middle.json"))
        content_files = list(output_path.glob("**/*_content_list.json"))

        top_md = [f for f in md_files if f.parent == output_path]
        sub_md = [f for f in md_files if f.parent != output_path]
        best_md = (top_md or sub_md)
        md_file = best_md[0] if best_md else None
        json_file = json_files[0] if json_files else None
        content_file = content_files[0] if content_files else None

        markdown = md_file.read_text(encoding="utf-8") if md_file else ""

        middle_json = {}
        if json_file:
            try:
                middle_json = json.loads(json_file.read_text(encoding="utf-8"))
            except Exception:
                pass

        content_list = []
        if content_file:
            try:
                content_list = json.loads(content_file.read_text(encoding="utf-8"))
            except Exception:
                pass

        blocks = self._extract_blocks(content_list, middle_json)
        from .layout_reader_engine import detect_columns_and_order, blocks_to_markdown
        blocks = detect_columns_and_order(blocks)
        markdown = blocks_to_markdown(blocks)

        return {
            "markdown": markdown,
            "blocks": blocks,
            "middle_json": middle_json,
            "images_dir": str(output_path / "images") if (output_path / "images").exists()
                          else str(next(output_path.glob("**/images"), Path("."))),
            "metadata": {
                "engine": "RapidDoc",
                "layout_model": "PP-DocLayoutV2",
                "formula_model": "PP-FormulaNet_plus",
                "ocr_engine": "RapidOCR",
                "table_model": "SLANET/UNITABLE",
            },
        }

    # ─── Block 提取 ──────────────────────────────────────────────────────

    def _extract_blocks(self, content_list: list, middle_json: dict) -> list[dict]:
        """将 content_list 转为标准化 block"""
        blocks = []
        if not content_list:
            return blocks

        for item in content_list:
            raw_type = item.get("type", "text")
            level = item.get("text_level", 0) or item.get("level", 0)
            text = item.get("text", "") or ""

            # 跳过丢弃项（页眉页脚）
            if raw_type == "discarded":
                continue

            # 类型映射
            if raw_type == "equation":
                btype = "formula"
                text = clean_latex(text)
            elif raw_type == "table":
                btype = "table"
                html_body = item.get("table_body", "")
                text = _html_table_to_md(html_body) if html_body else text
            elif raw_type == "image":
                btype = "image"
                text = item.get("img_path", "") or text
            elif level >= 1:
                btype = "heading"
                level = infer_heading_level(text, level)
            else:
                btype = "text"

            blocks.append({
                "type": btype,
                "text": text,
                "bbox": item.get("bbox", None),
                "page": item.get("page_idx", 0),
                "level": level if btype == "heading" else 0,
            })

        return blocks



# ─── HTML 表格 → Markdown ──────────────────────────────────────────────


class _TableParser(HTMLParser):
    """解析 HTML <table> → Markdown"""

    def __init__(self):
        super().__init__()
        self.rows = []
        self._cur_row = []
        self._cur_cell = ""
        self._in_td = False

    def handle_starttag(self, tag, attrs):
        if tag in ("td", "th"):
            self._in_td = True
            self._cur_cell = ""
        elif tag == "tr":
            self._cur_row = []

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self._in_td = False
            self._cur_row.append(self._cur_cell.strip())
        elif tag == "tr":
            if self._cur_row:
                self.rows.append(self._cur_row)

    def handle_data(self, data):
        if self._in_td:
            self._cur_cell += data


def _html_table_to_md(html: str) -> str:
    """将 HTML 表格转为 Markdown 表格"""
    if not html or "<table>" not in html:
        return ""

    parser = _TableParser()
    try:
        parser.feed(html)
    except Exception:
        return ""

    rows = parser.rows
    if not rows:
        return ""

    # 清理空格过多的单元格
    cleaned = []
    for row in rows:
        cleaned.append([clean_latex(cell) for cell in row])

    # 补全不等宽的行
    max_cols = max(len(r) for r in cleaned)
    for r in cleaned:
        while len(r) < max_cols:
            r.append("")

    # 生成 Markdown
    lines = []
    lines.append("| " + " | ".join(cleaned[0]) + " |")
    lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    for row in cleaned[1:]:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)
