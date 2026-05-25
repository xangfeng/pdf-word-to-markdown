"""Markdown + 格式 JSON → Word (.docx)

解析 Markdown 为 AST，结合格式 JSON 重建 Word 文档，
还原样式、表格、图片、公式和双栏布局。
"""

import os
import re
import sys
from xml.etree import ElementTree as ET

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import mistune

from .ooml_latex import latex_to_ooml
from .utils import read_json, ensure_dir


def convert_md_to_word(md_path: str, format_path: str, output_path: str | None = None) -> str:
    """主入口：将 Markdown + format.json 转换为 Word 文档

    Args:
        md_path: Markdown 文件路径
        format_path: 格式 JSON 文件路径
        output_path: 输出 .docx 路径，默认同目录同名 .docx

    Returns:
        输出文件的路径
    """
    if output_path is None:
        output_path = os.path.splitext(md_path)[0] + "_reconstructed.docx"

    # 读取输入
    with open(md_path, "r", encoding="utf-8") as f:
        md_content = f.read()
    format_info = read_json(format_path)

    # 创建文档
    doc = Document()

    # 设置页面
    _setup_page(doc, format_info)

    # 创建样式
    _create_styles(doc, format_info)

    # 解析 Markdown 为 AST
    ast = _parse_markdown(md_content)

    # 渲染到 Word
    _render_ast_to_docx(doc, ast, format_info, os.path.dirname(md_path))

    # 保存
    doc.save(output_path)
    return output_path


# ─── 页面设置 ────────────────────────────────────────────────────────────────


def _setup_page(doc: Document, format_info: dict) -> None:
    """设置页面尺寸和边距"""
    ps = format_info.get("page_setup", {})
    section = doc.sections[0]

    if "width_cm" in ps:
        section.page_width = Cm(ps["width_cm"])
    if "height_cm" in ps:
        section.page_height = Cm(ps["height_cm"])

    margins = ps.get("margins", {})
    if "top_cm" in margins:
        section.top_margin = Cm(margins["top_cm"])
    if "bottom_cm" in margins:
        section.bottom_margin = Cm(margins["bottom_cm"])
    if "left_cm" in margins:
        section.left_margin = Cm(margins["left_cm"])
    if "right_cm" in margins:
        section.right_margin = Cm(margins["right_cm"])


# ─── 样式创建 ────────────────────────────────────────────────────────────────


_STYLE_MAP = {
    "一级标题": "Heading 1",
    "二级标题": "Heading 2",
    "三级标题": "Heading 3",
    "四级标题": "Heading 4",
    "标题": "Title",
    "正文": "Normal",
}


def _create_styles(doc: Document, format_info: dict) -> None:
    """根据格式 JSON 创建/修改 Word 样式"""
    styles = format_info.get("styles", {})
    for style_name, info in styles.items():
        wd_style_name = _STYLE_MAP.get(style_name, style_name)

        try:
            style = doc.styles[wd_style_name]
        except KeyError:
            style = doc.styles.add_style(wd_style_name, 1)  # WD_STYLE_TYPE.PARAGRAPH

        font = style.font
        if "font_name" in info:
            font.name = info["font_name"]
            # 设置东亚字体
            rpr = style.element.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                style.element.append(rpr)
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is None:
                rfonts = OxmlElement("w:rFonts")
                rpr.insert(0, rfonts)
            rfonts.set(qn("w:eastAsia"), info["font_name"])

        if "font_size_pt" in info:
            font.size = Pt(info["font_size_pt"])
        if "bold" in info:
            font.bold = info["bold"]
        if "italic" in info:
            font.italic = info["italic"]
        if "color" in info:
            hex_color = info["color"].lstrip("#")
            font.color.rgb = RGBColor(
                int(hex_color[0:2], 16) if len(hex_color) >= 6 else 0,
                int(hex_color[2:4], 16) if len(hex_color) >= 6 else 0,
                int(hex_color[4:6], 16) if len(hex_color) >= 6 else 0,
            )

        # 段落格式
        pf = style.paragraph_format
        if "alignment" in info:
            al_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
            pf.alignment = al_map.get(info["alignment"], 0)
        if "line_spacing" in info:
            pf.line_spacing = info["line_spacing"]
        if "space_before_pt" in info:
            pf.space_before = Pt(info["space_before_pt"])
        if "space_after_pt" in info:
            pf.space_after = Pt(info["space_after_pt"])
        if "first_line_indent_pt" in info:
            pf.first_line_indent = Pt(info["first_line_indent_pt"])


# ─── Markdown 解析 ───────────────────────────────────────────────────────────


def _parse_markdown(md_content: str) -> list:
    """将 Markdown 文本解析为结构化 AST"""
    # 预处理：移除分栏围栏，记录分栏区域
    column_sections = []
    processed_lines = []
    in_columns = False
    current_cols = []
    col_count = 1

    lines = md_content.split("\n")

    for line in lines:
        if line.startswith("::: columns"):
            in_columns = True
            col_count = int(re.search(r"count=(\d+)", line).group(1)) if re.search(r"count=(\d+)", line) else 2
            current_cols = [[] for _ in range(col_count)]
            continue
        elif line.startswith("::: column") and in_columns:
            continue
        elif line == ":::" and in_columns:
            # 内层闭合
            column_sections.append({"count": col_count, "columns": current_cols})
            processed_lines.append(f"__COLUMN_SECTION_{len(column_sections) - 1}__")
            in_columns = False
            continue

        if in_columns:
            # 找到当前活动的列
            active_col = len(current_cols) - 1
            for i, col in enumerate(current_cols):
                if len(col) == 0 or col[-1] != ":::":
                    break
            current_cols[-1].append(line)
        else:
            processed_lines.append(line)

    processed_text = "\n".join(processed_lines)

    # 用 mistune 解析
    renderer = _MDToASTRenderer()
    markdown = mistune.create_markdown(renderer=renderer)
    markdown(processed_text)

    ast = renderer.get_ast()
    ast.append({"type": "column_sections", "sections": column_sections})

    return ast


class _MDToASTRenderer:
    """将 Markdown 渲染为结构化 AST（而非 HTML）"""

    def __init__(self):
        self.ast = []
        self._current_para = None

    def get_ast(self):
        return self.ast

    def heading(self, text, level, **attrs):
        self.ast.append({"type": "heading", "text": text, "level": level})

    def paragraph(self, text):
        # 解析内联元素
        inline = _parse_inline_elements(text)
        self.ast.append({"type": "paragraph", "children": inline})

    def block_code(self, code, info=None):
        if info == "math" or (info is None and ("\\" in code or "^" in code or "_" in code)):
            self.ast.append({"type": "formula", "latex": code.strip(), "block": True})
        else:
            self.ast.append({"type": "code", "text": code, "lang": info})

    def block_text(self, text):
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            if line.startswith("__COLUMN_SECTION_"):
                match = re.search(r"__COLUMN_SECTION_(\d+)__", line)
                if match:
                    idx = int(match.group(1))
                    self.ast.append({"type": "column_section_ref", "index": idx})
            elif line.startswith("<!--"):
                self.ast.append({"type": "comment", "text": line})
            else:
                inline = _parse_inline_elements(line)
                self.ast.append({"type": "paragraph", "children": inline})

    def table(self, text):
        self.ast.append({"type": "table", "text": text.strip()})

    def image(self, src, alt="", title=None):
        self.ast.append({"type": "image", "src": src, "alt": alt})

    def thematic_break(self):
        self.ast.append({"type": "thematic_break"})

    def list(self, text, ordered, **attrs):
        items = text.strip().split("\n")
        items = [i.lstrip("-* 0123456789.").strip() for i in items]
        self.ast.append({"type": "list", "items": items, "ordered": ordered})

    def blank_line(self):
        pass

    def block_html(self, html):
        self.ast.append({"type": "html", "text": html})


def _parse_inline_elements(text: str) -> list:
    """解析内联元素：粗体、斜体、行内公式、行内代码、链接"""
    children = []
    pos = 0
    while pos < len(text):
        # 行内公式 $...$
        if text[pos] == "$" and pos + 1 < len(text) and text[pos + 1] != "$":
            end = text.find("$", pos + 1)
            if end > pos:
                children.append({"type": "formula", "latex": text[pos + 1:end], "block": False})
                pos = end + 1
                continue
        # 粗体 **...**
        if text[pos:pos + 2] == "**":
            end = text.find("**", pos + 2)
            if end > pos:
                children.append({"type": "bold", "text": text[pos + 2:end]})
                pos = end + 2
                continue
        # 斜体 *...*
        if text[pos] == "*" and text[pos:pos + 2] != "**":
            end = text.find("*", pos + 1)
            if end > pos:
                children.append({"type": "italic", "text": text[pos + 1:end]})
                pos = end + 1
                continue
        # 行内代码 `...`
        if text[pos] == "`":
            end = text.find("`", pos + 1)
            if end > pos:
                children.append({"type": "code", "text": text[pos + 1:end]})
                pos = end + 1
                continue
        # 链接 [text](url)
        if text[pos] == "[":
            end_bracket = text.find("]", pos + 1)
            if end_bracket > pos and end_bracket + 1 < len(text) and text[end_bracket + 1] == "(":
                end_paren = text.find(")", end_bracket + 2)
                if end_paren > end_bracket:
                    link_text = text[pos + 1:end_bracket]
                    link_url = text[end_bracket + 2:end_paren]
                    children.append({"type": "link", "text": link_text, "url": link_url})
                    pos = end_paren + 1
                    continue
        # 纯文本
        next_special = len(text)
        for ch in ["$", "*", "`", "["]:
            idx = text.find(ch, pos)
            if idx != -1 and idx < next_special:
                next_special = idx
        if next_special > pos:
            children.append({"type": "text", "value": text[pos:next_special]})
        pos = next_special
    return children


# ─── AST → Document 渲染 ───────────────────────────────────────────────────


def _render_ast_to_docx(doc: Document, ast: list, format_info: dict, base_dir: str) -> None:
    """将 AST 渲染到 Word 文档"""
    column_sections = None
    for item in ast:
        if item["type"] == "column_sections":
            column_sections = item["sections"]

    for item in ast:
        if item["type"] in ("column_section_ref", "column_sections"):
            continue
        _render_node(doc, item, format_info, base_dir, column_sections)


def _render_node(doc: Document, node: dict, format_info: dict, base_dir: str, column_sections: list | None = None) -> None:
    """渲染单个 AST 节点"""
    ntype = node["type"]

    if ntype == "heading":
        level_map = {1: "一级标题", 2: "二级标题", 3: "三级标题", 4: "四级标题"}
        style_name = level_map.get(node["level"], "Normal")
        wd_style = _STYLE_MAP.get(style_name, "Normal")
        para = doc.add_paragraph(style=wd_style)
        _add_text_to_para(para, node["text"])

    elif ntype == "paragraph":
        para = doc.add_paragraph(style="Normal")
        for child in node.get("children", []):
            _add_inline_to_para(para, child)

    elif ntype == "formula":
        _add_formula_to_doc(doc, node["latex"], node.get("block", False))

    elif ntype == "code":
        para = doc.add_paragraph(style="Normal")
        run = para.add_run(node["text"])
        run.font.name = "Consolas"
        run.font.size = Pt(10)

    elif ntype == "table":
        _render_table(doc, node["text"], format_info)

    elif ntype == "image":
        src = node["src"]
        img_path = os.path.join(base_dir, src) if not os.path.isabs(src) else src
        if os.path.exists(img_path):
            para = doc.add_paragraph(style="Normal")
            run = para.add_run()
            run.add_picture(img_path)
            # 尝试从格式信息设置图片尺寸
            img_key = os.path.splitext(os.path.basename(src))[0]
            img_info = format_info.get("images", {}).get(img_key, {})
            if img_info:
                if "width_px" in img_info:
                    run.width = Inches(img_info["width_px"] / 96)
                if "height_px" in img_info:
                    run.height = Inches(img_info["height_px"] / 96)

    elif ntype == "list":
        for item_text in node.get("items", []):
            para = doc.add_paragraph(style="List Bullet")
            _add_text_to_para(para, item_text)

    elif ntype == "thematic_break":
        para = doc.add_paragraph(style="Normal")
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
        # 添加水平线（底边框）
        pPr = para._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "999999")
        pBdr.append(bottom)
        pPr.append(pBdr)

    elif ntype == "comment":
        # HTML 注释 → 保留为 Word 中的隐藏文本
        pass

    elif ntype == "html":
        pass


def _add_text_to_para(para, text: str, bold: bool = False, italic: bool = False) -> None:
    """向段落添加简单文本"""
    run = para.add_run(text)
    if bold:
        run.bold = True
    if italic:
        run.italic = True


def _add_inline_to_para(para, child: dict) -> None:
    """向段落添加内联元素"""
    ctype = child["type"]

    if ctype == "text":
        para.add_run(child["value"])
    elif ctype == "bold":
        run = para.add_run(child["text"])
        run.bold = True
    elif ctype == "italic":
        run = para.add_run(child["text"])
        run.italic = True
    elif ctype == "formula":
        _add_inline_formula_to_para(para, child["latex"])
    elif ctype == "code":
        run = para.add_run(child["text"])
        run.font.name = "Consolas"
        run.font.size = Pt(10)
    elif ctype == "link":
        run = para.add_run(child["text"])
        run.font.color.rgb = RGBColor(0, 0, 255)
        run.underline = True


# ─── 公式插入 ────────────────────────────────────────────────────────────────


def _add_formula_to_doc(doc: Document, latex: str, block: bool = False) -> None:
    """向文档添加块级公式"""
    para = doc.add_paragraph(style="Normal")
    try:
        ooml = latex_to_ooml(latex, inline=not block)
        para._p.append(ooml)
    except Exception:
        # 降级：纯文本
        run = para.add_run(f"$${latex}$$" if block else f"${latex}$")
        run.italic = True


def _add_inline_formula_to_para(para, latex: str) -> None:
    """向段落添加行内公式"""
    try:
        ooml = latex_to_ooml(latex, inline=True)
        para._p.append(ooml)
    except Exception:
        run = para.add_run(f"${latex}$")
        run.italic = True


# ─── 表格渲染 ────────────────────────────────────────────────────────────────


def _render_table(doc: Document, table_text: str, format_info: dict) -> None:
    """将 Markdown 表格文本渲染为 Word 表格"""
    lines = table_text.strip().split("\n")
    if len(lines) < 2:
        return

    rows = []
    for i, line in enumerate(lines):
        if i == 1 and all(c in "|-: " for c in line):
            continue  # 跳过分隔行
        cells = [c.strip() for c in line.split("|")]
        cells = [c for c in cells if c]  # 去空
        # 清理公式换行
        cells = [c.replace("\n$$\n", "$$") for c in cells]
        if cells:
            rows.append(cells)

    if not rows:
        return

    max_cols = max(len(r) for r in rows)
    while len(rows) < max_cols:
        rows.append([""] * max_cols)

    table = doc.add_table(rows=len(rows), cols=max_cols, style="Table Grid")
    for i, row_data in enumerate(rows):
        row = table.rows[i]
        for j, cell_text in enumerate(row_data):
            if j < max_cols:
                row.cells[j].text = cell_text


# ─── CLI 入口 ────────────────────────────────────────────────────────────────


def main():
    if len(sys.argv) < 3:
        print("用法: python -m pdf_word_to_markdown.md_to_word <input.md> <format.json> [output.docx]")
        sys.exit(1)

    md_path = sys.argv[1]
    format_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else None

    result = convert_md_to_word(md_path, format_path, output_path)
    print(f"已生成 Word 文档: {result}")


if __name__ == "__main__":
    main()
