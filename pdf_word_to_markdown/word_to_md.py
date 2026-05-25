"""Word (.docx) → Markdown + 格式 JSON

遍历 Word 文档的段落、表格、图片、公式，生成 Markdown 文件，
同时记录所有格式信息到 format.json，确保后续可完整还原。
"""

import os
import sys
from xml.etree import ElementTree as ET
from docx import Document
from docx.oxml.ns import qn
from .ooml_latex import ooml_to_latex
from .utils import (
    write_json, ensure_dir, align_str,
    emu_to_px, rgb_to_hex, extract_images_from_docx,
    clean_latex, infer_heading_level,
)

# OMML 命名空间
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def convert_word_to_md(docx_path: str, output_dir: str | None = None) -> dict:
    """主入口：将 .docx 文件转换为 Markdown + format.json

    Args:
        docx_path: Word 文档路径
        output_dir: 输出目录，默认为文档同目录下的 {name}_output/

    Returns:
        {"md": "markdown文件路径", "format": "格式json路径", "images": "图片目录路径"}
    """
    base = os.path.splitext(os.path.basename(docx_path))[0]
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(docx_path) or ".", f"{base}_output")
    ensure_dir(output_dir)

    images_dir = os.path.join(output_dir, "images")
    ensure_dir(images_dir)

    doc = Document(docx_path)

    # 提取图片
    image_map = extract_images_from_docx(doc, images_dir)

    # 格式信息收集
    format_info = {
        "meta": {"source": "docx", "source_file": os.path.basename(docx_path)},
        "page_setup": {},
        "styles": {},
        "columns": [],
        "tables": {},
        "images": {},
    }

    # 提取页面设置
    _extract_page_setup(doc, format_info)

    # 检测分栏
    column_sections = _detect_columns(doc)

    # 生成 Markdown 内容
    md_lines = []
    para_index = 0
    table_index = 0
    image_index = 0

    body = doc.element.body

    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            md_text, style_info = _process_paragraph(child, doc, image_map, images_dir, image_index)
            if md_text:
                md_lines.append(md_text)
                if style_info:
                    format_info["styles"].setdefault(style_info["style_name"], style_info)
            para_index += 1

        elif tag == "tbl":
            md_table, table_info = _process_table(child)
            if md_table:
                md_lines.append("")
                md_lines.append(md_table)
                md_lines.append("")
                if table_info:
                    format_info["tables"][f"table_{table_index}"] = table_info
                table_index += 1

        elif tag == "sdt":
            # 结构化文档标签（如目录），递归处理内部内容
            for inner in child:
                inner_tag = inner.tag.split("}")[-1] if "}" in inner.tag else inner.tag
                if inner_tag == "sdtContent":
                    for content_child in inner:
                        content_tag = content_child.tag.split("}")[-1] if "}" in content_child.tag else content_child.tag
                        if content_tag == "p":
                            md_text, style_info = _process_paragraph(content_child, doc, image_map, images_dir, image_index)
                            if md_text:
                                md_lines.append(md_text)
                                if style_info:
                                    format_info["styles"].setdefault(style_info["style_name"], style_info)
                            para_index += 1
                        elif content_tag == "tbl":
                            md_table, table_info = _process_table(content_child)
                            if md_table:
                                md_lines.append("")
                                md_lines.append(md_table)
                                md_lines.append("")
                                if table_info:
                                    format_info["tables"][f"table_{table_index}"] = table_info
                                table_index += 1

    # 写入 Markdown 文件
    md_path = os.path.join(output_dir, f"{base}.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    # 写入格式 JSON
    format_path = os.path.join(output_dir, f"{base}_format.json")
    write_json(format_info, format_path)

    return {"md": md_path, "format": format_path, "images": images_dir}


# ─── 页面设置 ────────────────────────────────────────────────────────────────


def _extract_page_setup(doc, format_info: dict) -> None:
    """提取页面尺寸和边距"""
    for section in doc.sections:
        ps = format_info["page_setup"]
        ps["width_cm"] = round(section.page_width / 360000, 2) if section.page_width else 21.0
        ps["height_cm"] = round(section.page_height / 360000, 2) if section.page_height else 29.7
        ps["margins"] = {
            "top_cm": round(section.top_margin / 360000, 2) if section.top_margin else 2.54,
            "bottom_cm": round(section.bottom_margin / 360000, 2) if section.bottom_margin else 2.54,
            "left_cm": round(section.left_margin / 360000, 2) if section.left_margin else 3.18,
            "right_cm": round(section.right_margin / 360000, 2) if section.right_margin else 3.18,
        }
        break  # 只取第一个节


# ─── 分栏检测 ────────────────────────────────────────────────────────────────


def _detect_columns(doc) -> list[dict]:
    """从节的 XML 中检测分栏设置"""
    columns = []
    for section in doc.sections:
        sect_pr = section._sectPr
        cols = sect_pr.find(qn("w:cols"))
        if cols is not None:
            num = int(cols.get(qn("w:num"), "1"))
            sep = cols.get(qn("w:sep"), "0") == "1"
            columns.append({"count": num, "separator": sep})
    return columns


# ─── 段落处理 ────────────────────────────────────────────────────────────────


def _process_paragraph(para_xml, doc, image_map: dict, images_dir: str, img_idx: int) -> tuple[str, dict | None]:
    """处理一个段落，返回 (Markdown文本, 样式信息)"""
    # 检查是否包含公式
    omaths = para_xml.findall(f"{{{M_NS}}}oMath") + para_xml.findall(f"{{{M_NS}}}oMathPara")

    # 获取段落样式
    style_name = "正文"
    para_props = para_xml.find(qn("w:pPr"))
    if para_props is not None:
        pstyle = para_props.find(qn("w:pStyle"))
        if pstyle is not None:
            style_name = pstyle.get(qn("w:val"), "正文")

    # 收集格式信息
    style_info = _extract_paragraph_style(para_xml, style_name)

    # 提取文本和公式
    parts = []
    _parse_paragraph_runs(para_xml, parts, image_map, images_dir, img_idx)

    if not parts:
        return "", None

    # 构建 Markdown（先获取文本用于层级推断）
    heading_level = _infer_heading_level_from_style(style_name)
    md = _build_markdown_from_parts(parts, style_name, heading_level)

    # 显示名
    heading_map = {
        "Heading 1": "一级标题", "heading 1": "一级标题",
        "Heading 2": "二级标题", "heading 2": "二级标题",
        "Heading 3": "三级标题", "heading 3": "三级标题",
        "Heading 4": "四级标题", "heading 4": "四级标题",
        "1 Heading 1": "一级标题", "2 Heading 2": "二级标题",
        "3 Heading 3": "三级标题", "Title": "标题",
    }
    display_name = heading_map.get(style_name, style_name)
    if display_name != style_name or heading_level > 0:
        style_info["style_name"] = display_name
        style_info["level"] = heading_level

    return md, style_info


def _parse_paragraph_runs(para_xml, parts: list, image_map: dict, images_dir: str, img_idx: int) -> None:
    """解析段落内的 runs（文本、公式、图片、超链接）"""
    for child in para_xml:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "r":
            # 普通文本 run
            text = ""
            bold = False
            italic = False
            for r_child in child:
                r_tag = r_child.tag.split("}")[-1] if "}" in r_child.tag else r_child.tag
                if r_tag == "t":
                    text += r_child.text or ""
                elif r_tag == "rPr":
                    bold = r_child.find(qn("w:b")) is not None
                    italic = r_child.find(qn("w:i")) is not None
                elif r_tag == "drawing":
                    _handle_drawing(r_child, parts, image_map)
                elif r_tag == "pict":
                    _handle_pict(r_child, parts, image_map)

            if text.strip():
                prefix = suffix = ""
                if bold:
                    prefix += "**"
                    suffix = "**" + suffix
                if italic:
                    prefix += "*"
                    suffix = "*" + suffix
                parts.append({"type": "text", "value": text.strip(), "prefix": prefix, "suffix": suffix})

        elif tag == "oMath" or tag == "oMathPara":
            latex = ooml_to_latex(child)
            is_block = (tag == "oMathPara")
            parts.append({"type": "formula", "value": latex, "block": is_block})

        elif tag == "hyperlink":
            # 超链接
            for h_child in child:
                h_tag = h_child.tag.split("}")[-1] if "}" in h_child.tag else h_child.tag
                if h_tag == "r":
                    t_elem = h_child.find(qn("w:t"))
                    if t_elem is not None:
                        text = t_elem.text or ""
                        parts.append({"type": "text", "value": text})

        elif tag == "drawing":
            _handle_drawing(child, parts, image_map)

        elif tag == "pict":
            _handle_pict(child, parts, image_map)


def _handle_drawing(drawing_elem, parts: list, image_map: dict) -> None:
    """处理 w:drawing 中的图片"""
    blip = drawing_elem.find(f".//{{{'http://schemas.openxmlformats.org/drawingml/2006/main'}}}blip")
    if blip is not None:
        embed = blip.get(f"{{{'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}}}embed")
        if embed and embed in image_map:
            img_path = image_map[embed]
            rel_path = os.path.join("images", os.path.basename(img_path))
            parts.append({"type": "image", "value": rel_path, "alt": "图片"})

        # 尝试获取图片尺寸
        ext = drawing_elem.find(f".//{{{'http://schemas.openxmlformats.org/drawingml/2006/main'}}}ext")
        if ext is not None:
            cx = ext.get("cx")
            cy = ext.get("cy")
            if cx and cy:
                width = int(cx)
                height = int(cy)
                if parts and parts[-1]["type"] == "image":
                    parts[-1]["width_px"] = emu_to_px(width)
                    parts[-1]["height_px"] = emu_to_px(height)


def _handle_pict(pict_elem, parts: list, image_map: dict) -> None:
    """处理 w:pict (VML 图片) 中的图片"""
    imagedata = pict_elem.find(f".//{{{'urn:schemas-microsoft-com:vml'}}}imagedata")
    if imagedata is not None:
        rid = imagedata.get(f"{{{'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}}}id")
        if rid and rid in image_map:
            img_path = image_map[rid]
            rel_path = os.path.join("images", os.path.basename(img_path))
            parts.append({"type": "image", "value": rel_path, "alt": "图片"})


def _build_markdown_from_parts(parts: list, style_name: str, heading_level: int = 0) -> str:
    """将 parsed parts 拼接为 Markdown 字符串"""
    # 根据层级生成前缀
    if heading_level >= 1:
        prefix = "#" * min(heading_level, 4) + " "
    else:
        heading_map = {
            "Heading 1": "# ", "heading 1": "# ",
            "Heading 2": "## ", "heading 2": "## ",
            "Heading 3": "### ", "heading 3": "### ",
            "Heading 4": "#### ", "heading 4": "#### ",
            "1 Heading 1": "# ", "2 Heading 2": "## ",
            "3 Heading 3": "### ", "Title": "# ",
        }
        prefix = heading_map.get(style_name, "")

    result = []
    for part in parts:
        if part["type"] == "text":
            text = part["value"]
            if part.get("prefix"):
                text = part["prefix"] + text + part.get("suffix", "")
            result.append(text)
        elif part["type"] == "formula":
            latex = clean_latex(part["value"])
            if part.get("block"):
                result.append(f"\n$$\n{latex}\n$$\n")
            else:
                result.append(f"${latex}$")
        elif part["type"] == "image":
            result.append(f"![{part.get('alt', '图片')}]({part['value']})")

    combined = " ".join([r for r in result if r.strip()])
    return prefix + combined if prefix else combined.strip()


def _infer_heading_level_from_style(style_name: str) -> int:
    """从 Word 样式名推断标题层级"""
    mapping = {
        "Heading 1": 1, "heading 1": 1, "1 Heading 1": 1,
        "Heading 2": 2, "heading 2": 2, "2 Heading 2": 2,
        "Heading 3": 3, "heading 3": 3, "3 Heading 3": 3,
        "Heading 4": 4, "heading 4": 4,
        "Title": 1,
    }
    return mapping.get(style_name, 0)


def _extract_paragraph_style(para_xml, style_name: str) -> dict:
    """从段落 XML 中提取格式信息"""
    info = {"style_name": style_name}

    para_props = para_xml.find(qn("w:pPr"))
    if para_props is not None:
        # 对齐
        jc = para_props.find(qn("w:jc"))
        if jc is not None:
            info["alignment"] = align_str(
                {"left": 0, "center": 1, "right": 2, "both": 3}.get(
                    jc.get(qn("w:val"), "left"), 0
                )
            )
        # 行距
        spacing = para_props.find(qn("w:spacing"))
        if spacing is not None:
            line = spacing.get(qn("w:line"))
            if line:
                info["line_spacing"] = round(int(line) / 240, 2)
            before = spacing.get(qn("w:before"))
            if before:
                info["space_before_pt"] = round(int(before) / 20, 2)
            after = spacing.get(qn("w:after"))
            if after:
                info["space_after_pt"] = round(int(after) / 20, 2)
        # 首行缩进
        ind = para_props.find(qn("w:ind"))
        if ind is not None:
            first_line = ind.get(qn("w:firstLine"))
            if first_line:
                info["first_line_indent_pt"] = round(int(first_line) / 20, 2)

    # 第一个 run 的字体信息
    first_run = para_xml.find(qn("w:r"))
    if first_run is not None:
        run_props = first_run.find(qn("w:rPr"))
        if run_props is not None:
            rfonts = run_props.find(qn("w:rFonts"))
            if rfonts is not None:
                font = rfonts.get(qn("w:eastAsia")) or rfonts.get(qn("w:ascii")) or "宋体"
                info["font_name"] = font
            sz = run_props.find(qn("w:sz"))
            if sz is not None:
                info["font_size_pt"] = round(int(sz.get(qn("w:val"))) / 2, 1)
            color = run_props.find(qn("w:color"))
            if color is not None:
                val = color.get(qn("w:val"))
                if val and val != "auto":
                    info["color"] = f"#{val}"
            info["bold"] = run_props.find(qn("w:b")) is not None
            info["italic"] = run_props.find(qn("w:i")) is not None

    return info


# ─── 表格处理 ────────────────────────────────────────────────────────────────


def _process_table(tbl_xml) -> tuple[str, dict | None]:
    """处理表格 XML，返回 (Markdown表格, 表格格式信息)"""
    rows = tbl_xml.findall(f"{{{W_NS}}}tr")
    if not rows:
        return "", None

    # 获取表格属性
    tbl_pr = tbl_xml.find(f"{{{W_NS}}}tblPr")
    col_widths = []
    if tbl_pr is not None:
        tbl_grid = tbl_pr.find(f"{{{W_NS}}}tblGrid")
        if tbl_grid is not None:
            for gc in tbl_grid.findall(f"{{{W_NS}}}gridCol"):
                w = gc.get(f"{{{W_NS}}}w")
                if w:
                    col_widths.append(int(w))

    # 解析所有行
    table_data = []
    for row in rows:
        cells = row.findall(f"{{{W_NS}}}tc")
        row_data = []
        for cell in cells:
            text_parts = []
            for para in cell.findall(f"{{{W_NS}}}p"):
                cell_parts = []
                _parse_paragraph_runs(para, cell_parts, {}, "", 0)
                line = _build_markdown_from_parts(cell_parts, "")
                if line.strip():
                    text_parts.append(line.strip())
            # 清理公式换行
            cell_text = " ".join(text_parts)
            cell_text = cell_text.replace("\n$$\n", "$$").replace("\n", " ")
            row_data.append(cell_text)
        table_data.append(row_data)

    if not table_data:
        return "", None

    # 列数
    max_cols = max(len(r) for r in table_data)
    for r in table_data:
        while len(r) < max_cols:
            r.append("")

    # 生成 Markdown 表格
    md = _table_to_markdown(table_data)

    # 表格格式信息
    total_width = sum(col_widths) if col_widths else 0
    table_info = {
        "rows": len(table_data),
        "cols": max_cols,
    }
    if col_widths and total_width:
        table_info["col_widths"] = [round(w / total_width, 3) for w in col_widths[:max_cols]]
    # 检测是否有网格线
    tbl_borders = tbl_pr.find(f"{{{W_NS}}}tblBorders") if tbl_pr is not None else None
    table_info["bordered"] = tbl_borders is not None

    return md, table_info


def _table_to_markdown(data: list[list[str]]) -> str:
    """将二维表格数据转为 Markdown 表格字符串"""
    if not data:
        return ""

    lines = []
    # 表头
    lines.append("| " + " | ".join(data[0]) + " |")
    # 分隔行
    lines.append("| " + " | ".join(["---"] * len(data[0])) + " |")
    # 数据行
    for row in data[1:]:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)


# ─── CLI 入口 ────────────────────────────────────────────────────────────────


def main():
    if len(sys.argv) < 2:
        print("用法: python -m pdf_word_to_markdown.word_to_md <input.docx> [output_dir]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    result = convert_word_to_md(input_path, output_dir)
    print(f"Markdown: {result['md']}")
    print(f"格式文档: {result['format']}")
    print(f"图片目录: {result['images']}")


if __name__ == "__main__":
    main()
