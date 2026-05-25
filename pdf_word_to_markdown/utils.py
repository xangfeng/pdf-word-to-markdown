"""共用工具函数：文件读写、JSON处理、颜色转换、单位换算"""

import json
import os
import re
from pathlib import Path
from typing import Any


def read_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def write_json(data: dict, path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def ensure_dir(path: str) -> str:
    Path(path).mkdir(parents=True, exist_ok=True)
    return path


def sanitize_filename(name: str) -> str:
    """移除文件名中的非法字符"""
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def rgb_to_hex(rgb: tuple | None) -> str:
    """RGB元组转十六进制颜色字符串"""
    if rgb is None or len(rgb) < 3:
        return "#000000"
    return "#{:02X}{:02X}{:02X}".format(rgb[0], rgb[1], rgb[2])


def emu_to_px(emu: int) -> int:
    """EMU (English Metric Units) 转 像素 (96 DPI)"""
    return round(emu / 9525)


def emu_to_cm(emu: int) -> float:
    """EMU 转 厘米"""
    return round(emu / 360000, 4)


def extract_images_from_docx(doc, output_dir: str) -> dict[str, str]:
    """从 docx 文档中提取所有图片到输出目录"""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    image_map = {}
    ensure_dir(output_dir)

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image_data = rel.target_part.blob
            ext = rel.target_part.ext
            img_name = f"img_{rel.rId}.{ext}"
            img_path = os.path.join(output_dir, img_name)
            with open(img_path, "wb") as f:
                f.write(image_data)
            image_map[rel.rId] = img_path

    return image_map


def align_str(alignment: int | None) -> str:
    """python-docx 对齐常量 → 字符串"""
    mapping = {0: "left", 1: "center", 2: "right", 3: "justify"}
    return mapping.get(alignment, "left") if alignment is not None else "left"


def clean_latex(latex: str) -> str:
    """清理 LaTeX token 间多余空格

    \"s i g ( P _ { 1 } , P _ { 2 } )\" → \"sig(P_{1},P_{2})\"
    \"\\\\frac { a } { b }\" → \"\\\\frac{a}{b}\"
    """
    inner = latex.strip()
    inner = re.sub(r'^\$\$?\s*', '', inner)
    inner = re.sub(r'\s*\$\$?$', '', inner)
    # 去掉 LaTeX 特殊字符周围的空格
    inner = re.sub(r'\s*([_{}^{},\\\\])\s*', r'\1', inner)
    # 合并被空格拆开的单个字母/数字序列
    for _ in range(3):
        new_inner = re.sub(r'\b([a-zA-Z0-9]) ([a-zA-Z0-9])\b', r'\1\2', inner)
        if new_inner == inner:
            break
        inner = new_inner
    return inner.strip()


def infer_heading_level(text: str, base_level: int = 1) -> int:
    """从文本编号模式推断标题层级: 1→H1, 1.1→H2, 1.1.1→H3"""
    m = re.match(r'(\d+(?:\.\d+)*)\s*[　\s]', text)
    if m:
        return min(m.group(1).count(".") + 1, 4)
    m = re.match(r'(\d+(?:\.\d+)*)\.?\s+', text)
    if m:
        return min(m.group(1).count(".") + 1, 4)
    if re.match(r'第[一二三四五六七八九十\d]+章', text):
        return 1
    if re.match(r'第[一二三四五六七八九十\d]+节', text):
        return 2
    return base_level


