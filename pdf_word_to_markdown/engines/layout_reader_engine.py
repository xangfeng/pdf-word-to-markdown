"""分栏阅读顺序引擎

内置两套方案:
  1. heuristic (默认): 坐标聚类法，快速，无需模型
  2. layoutreader (可选): LayoutLMv3 阅读顺序模型，高精度

pdf_oxide ColumnAware 算法思路:
  - X 坐标聚类 → 识别栏数
  - 通栏块检测（宽度 > 单栏 1.5x）
  - 阅读顺序: 栏内从上到下，栏间从左到右
"""

import numpy as np
from collections import defaultdict


def detect_columns_and_order(blocks: list[dict]) -> list[dict]:
    """主入口：分析 blocks 的坐标，返回带栏号和排序的 blocks

    每个 block 需包含: bbox=[x0,y0,x1,y1], page
    增加: col (栏号 0/1/2...), span (是否通栏), order (阅读序号)
    """
    if not blocks:
        return blocks

    # 逐页处理
    pages = defaultdict(list)
    for b in blocks:
        pages[b.get("page", 0)].append(b)

    result = []
    order_counter = [0]  # 用列表以便在闭包中修改

    for page_num in sorted(pages.keys()):
        page_blocks = pages[page_num]
        page_w = _estimate_page_width(page_blocks)

        # 检测该页的栏数和栏边界
        col_count, col_ranges = _detect_columns_on_page(page_blocks, page_w)

        # 分配栏号 + 检测通栏
        for b in page_blocks:
            bbox = b.get("bbox", [0, 0, 100, 10])
            if len(bbox) == 4:
                bw = bbox[2] - bbox[0]
                cx = (bbox[0] + bbox[2]) / 2
                # 通栏检测: 宽度超过单栏 1.5x
                if col_count > 1 and bw > (page_w / col_count) * 1.3:
                    b["span"] = True
                    b["col"] = -1
                else:
                    b["span"] = False
                    b["col"] = _assign_column(cx, col_ranges)
            else:
                b["col"] = 0
                b["span"] = False

        # 栏内排序: y 从上到下
        col_groups = defaultdict(list)
        span_blocks = []
        for b in page_blocks:
            if b.get("span"):
                span_blocks.append(b)
            else:
                col_groups[b.get("col", 0)].append(b)

        for col_idx in range(col_count):
            col_groups[col_idx].sort(key=lambda b: b.get("bbox", [0, 0, 0, 0])[1])

        # 组装: 通栏块插入对应 y 位置
        ordered = _interleave_spans(col_groups, span_blocks, col_count)

        for b in ordered:
            b["order"] = order_counter[0]
            order_counter[0] += 1
            result.append(b)

    return result


def _estimate_page_width(blocks: list) -> float:
    """估算页面宽度"""
    rights = []
    for b in blocks:
        bbox = b.get("bbox", [])
        if len(bbox) == 4 and bbox[2] > 0:
            rights.append(bbox[2])
    return max(rights) if rights else 595


def _detect_columns_on_page(blocks: list, page_w: float) -> tuple[int, list]:
    """检测页面栏数: X 坐标聚类 + 间隙分析

    学术论文最多 2 栏，需要明显间隙(>8%页宽)才判定为多栏
    """
    x_centers = []
    widths = []
    for b in blocks:
        bbox = b.get("bbox", [])
        if len(bbox) == 4:
            x_centers.append((bbox[0] + bbox[2]) / 2)
            widths.append(bbox[2] - bbox[0])

    if len(x_centers) < 5:
        return 1, [(0, page_w)]

    xs = sorted(x_centers)
    gaps = [xs[i+1] - xs[i] for i in range(len(xs)-1)]
    if not gaps:
        return 1, [(0, page_w)]

    # 找显著间隙（>8% 页宽 = 约 48pt on 595pt page）
    min_gap = page_w * 0.08
    big_gaps = [(i, g) for i, g in enumerate(gaps) if g > min_gap]

    if not big_gaps:
        return 1, [(0, page_w)]

    # 取最大的 1 个间隙（双栏），2 个间隙（三栏）太罕见
    big_gaps.sort(key=lambda x: -x[1])
    best_gap = big_gaps[0]

    col_count = 2  # 学术论文上限 2 栏
    sep_x = xs[best_gap[0]] + best_gap[1] / 2

    return col_count, [(0, sep_x), (sep_x, page_w)]


def _assign_column(cx: float, col_ranges: list) -> int:
    """根据 x 中心分配栏号"""
    for i, (xmin, xmax) in enumerate(col_ranges):
        if xmin <= cx <= xmax:
            return i
    # 找最近的
    best = 0
    best_d = float("inf")
    for i, (xmin, xmax) in enumerate(col_ranges):
        mid = (xmin + xmax) / 2
        d = abs(cx - mid)
        if d < best_d:
            best_d = d
            best = i
    return best


def _interleave_spans(col_groups: dict, span_blocks: list, col_count: int):
    """将通栏块插入到正确 y 位置"""

    # 构建每栏的 y 范围
    col_y_ranges = {}
    for ci in range(col_count):
        cg = col_groups.get(ci, [])
        if cg:
            col_y_ranges[ci] = (
                cg[0].get("bbox", [0, 0, 0, 0])[1],
                cg[-1].get("bbox", [0, 0, 0, 0])[3]
            )

    # 按栏从左到右，栏内从上到下扁平化
    result = []
    span_used = set()

    for ci in range(col_count):
        cg = col_groups.get(ci, [])
        for b in cg:
            by = b.get("bbox", [0, 0, 0, 0])[1]

            # 插入位于此块之前的通栏块
            for si, sb in enumerate(span_blocks):
                if si in span_used:
                    continue
                sby = sb.get("bbox", [0, 0, 0, 0])[1]
                if sby < by:
                    result.append(sb)
                    span_used.add(si)

            result.append(b)

    # 剩余通栏块放在最后
    for si, sb in enumerate(span_blocks):
        if si not in span_used:
            result.append(sb)

    return result


# ─── Markdown 渲染（分栏感知）───────────────────────────────────────────────


def blocks_to_markdown(blocks: list[dict]) -> str:
    """Render ordered blocks as Markdown with column markers

    Page-by-page: detect columns, render per-column, handle span blocks
    """
    pages = defaultdict(list)
    for b in blocks:
        pages[b.get("page", 0)].append(b)

    all_lines = []
    for page_num in sorted(pages.keys()):
        page_blocks = pages[page_num]
        if not page_blocks:
            continue

        col_count = _page_col_count(page_blocks)
        lines = _render_page(page_blocks, col_count, page_num, len(pages))
        all_lines.extend(lines)

    return "\n".join(all_lines)


def _page_col_count(blocks: list) -> int:
    cols = set()
    for b in blocks:
        if not b.get("span", False):
            cols.add(b.get("col", 0))
    return max(len(cols), 1)


def _render_page(blocks: list, col_count: int, page_num: int, total_pages: int) -> list[str]:
    """渲染一页"""
    lines = []

    # 页标记（多页时）
    if total_pages > 1:
        page_w = max((b.get("bbox", [0, 0, 595, 0])[2] for b in blocks if len(b.get("bbox", [])) == 4), default=595)
        lines.append(f"<!-- 第 {page_num + 1} 页 -->")

    if col_count <= 1:
        # 单栏页
        for b in blocks:
            md = _block_to_md(b)
            if md:
                lines.append(md)
    else:
        # 多栏页
        col_groups = _group_by_column(blocks, col_count)
        spans = col_groups.pop("_spans", [])
        lines.append(f"::: columns count={col_count}")
        for ci in range(col_count):
            lines.append("::: column")
            for b in col_groups.get(ci, []):
                md = _block_to_md(b)
                if md:
                    lines.append(md)
            lines.append(":::")
        lines.append(":::")
        # 通栏块放在分栏之后
        for b in spans:
            md = _block_to_md(b)
            if md:
                lines.append(md)

    lines.append("")
    return lines


def _group_by_column(blocks: list, col_count: int) -> dict:
    """Group blocks by column, handling span blocks"""
    groups = defaultdict(list)
    for b in blocks:
        if b.get("span", False):
            continue
        col = min(b.get("col", 0), col_count - 1)
        groups[col].append(b)

    # 确保所有列都有 key
    result = {i: groups.get(i, []) for i in range(col_count)}

    # 通栏块放最后
    spans = [b for b in blocks if b.get("span", False)]
    if spans:
        result["_spans"] = spans

    return result


def _block_to_md(b: dict) -> str:
    """单个 block → Markdown"""
    t = b.get("type", "text")
    text = b.get("text", "") or ""
    if not text.strip() and t not in ("image",):
        return ""

    if t == "heading":
        level = b.get("level", 1)
        return f"{'#' * min(level, 4)} {text}"
    elif t == "formula":
        if text.startswith("$$"):
            return text
        # 块级公式（来自 equation 类型）
        return f"$$\n{text}\n$$"
    elif t == "table":
        return f"\n{text}\n"
    elif t == "image":
        img = b.get("img_path", "") or text
        if img and not img.startswith("images/"):
            img = f"images/{img}"
        return f"![]({img})" if img else ""
    else:
        # 正文: 清理行内公式空格
        return _clean_text(text)


def _clean_text(text: str) -> str:
    """清理行内公式 ($...$) 中的多余空格"""
    import re
    def clean_inline(m):
        latex = m.group(1)
        # 去掉 LaTeX 特殊字符周围空格
        latex = re.sub(r'\s*([_{}^{},\\])\s*', r'\1', latex)
        # 合并被拆开的字母
        for _ in range(3):
            new_latex = re.sub(r'\b([a-zA-Z0-9]) ([a-zA-Z0-9])\b', r'\1\2', latex)
            if new_latex == latex:
                break
            latex = new_latex
        return "$" + latex + "$"
    return re.sub(r'\$([^$]+)\$', clean_inline, text)
