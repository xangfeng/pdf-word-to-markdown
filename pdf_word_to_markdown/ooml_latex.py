"""OMML ↔ LaTeX 公式双向转换

OMML (Office Math Markup Language) 是 Word 中数学公式的 XML 表示。
本模块实现 OMML XML → LaTeX 的提取，以及 LaTeX → MathML → OMML 的生成。
"""

import re
from xml.etree import ElementTree as ET

# ─── OMML 命名空间 ───
NSMAP = {
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
for prefix, uri in NSMAP.items():
    ET.register_namespace(prefix, uri)


def _tag(name: str) -> str:
    return f"{{{NSMAP['m']}}}{name}"


# ─── OMML → LaTeX ──────────────────────────────────────────────────────────


def ooml_to_latex(elem) -> str:
    """将 OMML XML 元素递归转换为 LaTeX 字符串"""
    if elem is None:
        return ""

    tag_local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

    handlers = {
        "oMath": _handle_math,
        "oMathPara": _handle_math_para,
        "r": _handle_run,
        "f": _handle_fraction,
        "rad": _handle_radical,
        "sSup": _handle_sup,
        "sSub": _handle_sub,
        "sSubSup": _handle_subsup,
        "nary": _handle_nary,
        "acc": _handle_accent,
        "bar": _handle_bar,
        "d": _handle_delimiter,
        "groupChr": _handle_group_chr,
        "eqArr": _handle_eq_arr,
        "m": _handle_matrix,
        "limLow": _handle_lim_low,
        "limUpp": _handle_lim_upp,
        "borderBox": _handle_border_box,
        "box": _handle_box,
        "func": _handle_function,
        "phant": _handle_phantom,
        "spacing": _handle_phantom,
        "sPre": _handle_spre,
    }

    handler = handlers.get(tag_local, _handle_default)
    return handler(elem)


def _children_text(elem) -> str:
    """递归处理所有子元素，拼接文本"""
    parts = []
    for child in elem:
        part = ooml_to_latex(child)
        if part:
            parts.append(part)
    if not parts and elem.text:
        return elem.text.strip()
    return " ".join(parts)


def _handle_math(elem) -> str:
    return _children_text(elem)


def _handle_math_para(elem) -> str:
    return _children_text(elem)


def _handle_run(elem) -> str:
    """处理 m:r — 数学文本运行，提取纯文本或用 $$ 包裹"""
    parts = []
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "t":
            text = (child.text or "").strip()
            if text:
                parts.append(text)
        else:
            part = ooml_to_latex(child)
            if part:
                parts.append(part)
    return " ".join(parts)


def _handle_fraction(elem) -> str:
    """m:f → \\frac{分子}{分母}"""
    num = den = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        content = _children_text(child)
        if tag_local == "num":
            num = content
        elif tag_local == "den":
            den = content
    return f"\\frac{{{num}}}{{{den}}}"


def _handle_radical(elem) -> str:
    """m:rad → \\sqrt[deg]{base}"""
    deg = ""
    base = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "deg":
            deg = _children_text(child)
        elif tag_local == "e":
            base = _children_text(child)
    if deg:
        return f"\\sqrt[{deg}]{{{base}}}"
    return f"\\sqrt{{{base}}}"


def _handle_sup(elem) -> str:
    return f"^{{{_children_text(elem)}}}"


def _handle_sub(elem) -> str:
    return f"_{{{_children_text(elem)}}}"


def _handle_subsup(elem) -> str:
    sub = sup = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "sub":
            sub = _children_text(child)
        elif tag_local == "sup":
            sup = _children_text(child)
    return f"_{{{sub}}}^{{{sup}}}"


def _handle_nary(elem) -> str:
    """m:nary → \\sum, \\prod, \\int 等大型运算符"""
    nary_map = {
        "∑": "\\sum",
        "∏": "\\prod",
        "∐": "\\coprod",
        "∫": "\\int",
        "∬": "\\iint",
        "∭": "\\iiint",
        "∮": "\\oint",
        "⋀": "\\bigwedge",
        "⋁": "\\bigvee",
        "⋂": "\\bigcap",
        "⋃": "\\bigcup",
    }

    # 查找运算符字符
    char = ""
    sub = sup = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "chr":
            char = _children_text(child)
        elif tag_local == "sub":
            sub = _children_text(child)
        elif tag_local == "sup":
            sup = _children_text(child)

    latex = nary_map.get(char, char)
    if sub:
        latex += f"_{{{sub}}}"
    if sup:
        latex += f"^{{{sup}}}"
    return latex


def _handle_accent(elem) -> str:
    """m:acc → \\bar, \\hat, \\dot, \\ddot 等"""
    accent_map = {
        "̅": "\\bar",
        "̂": "\\hat",
        "̃": "\\tilde",
        "̇": "\\dot",
        "̈": "\\ddot",
        "⃗": "\\vec",
        "̆": "\\breve",
        "̌": "\\check",
        "̀": "\\grave",
        "́": "\\acute",
    }
    char = ""
    base = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "chr":
            char = _children_text(child)
        elif tag_local == "e":
            base = _children_text(child)
    latex = accent_map.get(char, "\\bar")
    return f"{latex}{{{base}}}"


def _handle_bar(elem) -> str:
    """m:bar → \\overline"""
    return f"\\overline{{{_children_text(elem)}}}"


def _handle_delimiter(elem) -> str:
    """m:d → 括号等定界符包裹的内容"""
    open_chars = ""
    close_chars = ""
    inner = ""

    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "dPr":
            for prop in child:
                ptag = prop.tag.split("}")[-1] if "}" in prop.tag else prop.tag
                if ptag == "begChr":
                    open_chars = (prop.text or "").strip()
                elif ptag == "endChr":
                    close_chars = (prop.text or "").strip()
        elif tag_local == "e":
            inner = _children_text(child)

    # 映射常见定界符
    delim_map = {
        "(": ("(", ")"),
        ")": ("(", ")"),
        "[": ("[", "]"),
        "|": ("|", "|"),
        "‖": ("\\|", "\\|"),
        "{": ("\\{", "\\}"),
        "⟦": ("\\llbracket", "\\rrbracket"),
        "⌊": ("\\lfloor", "\\rfloor"),
        "⌈": ("\\lceil", "\\rceil"),
        "⟨": ("\\langle", "\\rangle"),
    }

    left, right = delim_map.get(open_chars, (open_chars, close_chars))
    return f"\\left{left} {inner} \\right{right}"


def _handle_group_chr(elem) -> str:
    """m:groupChr → \\overbrace, \\underbrace 等"""
    char = ""
    inner = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "chr":
            char = _children_text(child)
        elif tag_local == "e":
            inner = _children_text(child)

    if "⏞" in char:
        return f"\\overbrace{{{inner}}}"
    elif "⏟" in char:
        return f"\\underbrace{{{inner}}}"
    return inner


def _handle_eq_arr(elem) -> str:
    """m:eqArr → \\begin{aligned}...\\end{aligned}"""
    lines = []
    for child in elem:
        if child.tag.split("}")[-1] == "e":
            lines.append(_children_text(child))
    inner = " \\\\ ".join(lines)
    return f"\\begin{{aligned}} {inner} \\end{{aligned}}"


def _handle_matrix(elem) -> str:
    """m:m → \\begin{matrix}...\\end{matrix}"""
    rows = []
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "mr":
            cells = []
            for cell in child:
                if cell.tag.split("}")[-1] == "e":
                    cells.append(_children_text(cell))
            rows.append(" & ".join(cells))
    inner = " \\\\ ".join(rows)
    return f"\\begin{{matrix}} {inner} \\end{{matrix}}"


def _handle_lim_low(elem) -> str:
    """m:limLow → \\lim_{}"""
    base = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "e":
            base = _children_text(child)
        elif tag_local == "lim":
            lim_val = _children_text(child)
            return f"{base}_{{{lim_val}}}"
    return base


def _handle_lim_upp(elem) -> str:
    """m:limUpp → \\lim^{}"""
    base = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "e":
            base = _children_text(child)
        elif tag_local == "lim":
            lim_val = _children_text(child)
            return f"{base}^{{{lim_val}}}"
    return base


def _handle_border_box(elem) -> str:
    """m:borderBox → \\boxed{}"""
    inner = ""
    for child in elem:
        if child.tag.split("}")[-1] == "e":
            inner = _children_text(child)
    return f"\\boxed{{{inner}}}"


def _handle_box(elem) -> str:
    return _children_text(elem)


def _handle_function(elem) -> str:
    """m:func → \\sin, \\cos, \\lim 等"""
    func_name = ""
    arg = ""
    for child in elem:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "fName":
            func_name = _children_text(child)
        elif tag_local == "e":
            arg = _children_text(child)
    return f"\\{func_name} {{{arg}}}"


def _handle_phantom(elem) -> str:
    return ""


def _handle_spre(elem) -> str:
    return _children_text(elem)


def _handle_default(elem) -> str:
    return _children_text(elem)


# ─── 从段落 XML 提取公式 ────────────────────────────────────────────────────


def extract_formulas_from_paragraph(para_xml) -> list[dict]:
    """从段落的 XML 中提取所有公式（OMML 或 MathML），返回类型和 LaTeX"""
    formulas = []
    root = para_xml

    for elem in root.iter():
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag in ("oMath", "oMathPara"):
            latex = ooml_to_latex(elem)
            formulas.append({"type": "block" if tag == "oMathPara" else "inline", "latex": latex})
        elif tag == "math":
            # MathML 公式（来自其他来源）
            latex = _mathml_to_latex(elem)
            formulas.append({"type": "inline", "latex": latex})

    return formulas


# ─── LaTeX → MathML → OMML ──────────────────────────────────────────────────


def latex_to_mathml(latex: str) -> str:
    """将 LaTeX 字符串转为 MathML，使用 latex2mathml 库"""
    from latex2mathml.converter import convert

    return convert(latex)


def mathml_to_ooml(mathml_str: str, inline: bool = True) -> ET.Element:
    """将 MathML 字符串转为 OMML XML 元素"""
    # 解析 MathML
    mathml_ns = "http://www.w3.org/1998/Math/MathML"
    root = ET.fromstring(mathml_str)

    tag = _tag("oMath") if inline else _tag("oMathPara")
    omath = ET.Element(tag)

    _mathml_elem_to_ooml(root, omath)

    return omath


# MathML 元素 → OMML 元素的映射表
_MATHML_TO_OMML_TAG = {
    "mfrac": "f",
    "msqrt": "rad",
    "mroot": "rad",
    "msup": "sSup",
    "msub": "sSub",
    "msubsup": "sSubSup",
    "mover": "groupChr",
    "munder": "limLow",
    "munderover": "limUpp",
    "mrow": "r",
    "mtable": "m",
    "mtr": "mr",
    "mtd": "e",
    "mfenced": "d",
    "mo": "r",
    "mi": "r",
    "mn": "r",
    "mtext": "r",
    "mspace": "r",
}


def _mathml_elem_to_ooml(ml_elem, om_parent) -> None:
    """递归将 MathML 子元素转为 OMML 结构"""
    ml_tag = ml_elem.tag.split("}")[-1] if "}" in ml_elem.tag else ml_elem.tag
    om_tag = _MATHML_TO_OMML_TAG.get(ml_tag, ml_tag)

    om_elem = ET.SubElement(om_parent, _tag(om_tag))

    # 处理文本内容
    if ml_elem.text and ml_elem.text.strip():
        t = ET.SubElement(om_elem, _tag("t"))
        t.text = ml_elem.text.strip()

    # 处理属性
    _convert_mathml_attrs(ml_elem, om_elem)

    # 递归处理子元素
    for child in ml_elem:
        _mathml_elem_to_ooml(child, om_elem)

    # 处理尾文本
    if ml_elem.tail and ml_elem.tail.strip():
        t = ET.SubElement(om_parent, _tag("t"))
        t.text = ml_elem.tail.strip()


def _convert_mathml_attrs(ml_elem, om_elem) -> None:
    """转换 MathML 元素的属性到 OMML"""
    ml_tag = ml_elem.tag.split("}")[-1] if "}" in ml_elem.tag else ml_elem.tag

    # 处理 mfrac 属性
    if ml_tag == "mfrac":
        linethick = ml_elem.get("linethickness", "1")
        if linethick != "1":
            f_pr = ET.Element(_tag("fPr"))
            ET.SubElement(f_pr, _tag("type")).set(_tag("val"), "noBar")
            om_elem.insert(0, f_pr)

    # 处理 mo 的 fence/stretchy 属性
    if ml_tag == "mo":
        fence = ml_elem.get("fence")
        stretchy = ml_elem.get("stretchy")
        char = (ml_elem.text or "").strip()

        # 括号 → d 元素
        if char in "()[]{}":
            # 重新标记为定界符
            d_pr = ET.Element(_tag("dPr"))
            beg = ET.SubElement(d_pr, _tag("begChr"))
            beg.set(_tag("val"), char)
            end_char = {"(": ")", "[": "]", "{": "}"}.get(char, char)
            end = ET.SubElement(d_pr, _tag("endChr"))
            end.set(_tag("val"), end_char)
            om_elem.insert(0, d_pr)


# ─── LaTeX → OMML（完整流程）────────────────────────────────────────────────


def latex_to_ooml(latex: str, inline: bool = True) -> ET.Element:
    """一站式：LaTeX 字符串 → OMML XML 元素"""
    try:
        mathml_str = latex_to_mathml(latex)
        return mathml_to_ooml(mathml_str, inline)
    except Exception:
        # 降级：创建简单文本运行
        fallback = ET.Element(_tag("oMath"))
        r = ET.SubElement(fallback, _tag("r"))
        t = ET.SubElement(r, _tag("t"))
        t.text = latex
        return fallback


def _mathml_to_latex(elem) -> str:
    """MathML 元素 → LaTeX 的简单转换（用于从其他来源的公式）"""
    # 简易实现：提取所有文本
    texts = []
    for t in elem.iter():
        if t.text and t.text.strip():
            texts.append(t.text.strip())
    return "".join(texts)
