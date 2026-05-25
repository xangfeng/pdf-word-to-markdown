"""Docling 引擎封装 —— IBM 文档解析

Docling (Apache 2.0) 核心能力：
  - DocLayNet (heron-101): 版面检测，78.0% mAP
  - TableFormer: 表格结构还原，cell-F1=0.89（表格最强！）
  - 阅读顺序恢复
  - 多格式输入（PDF, DOCX, PPTX, XLSX, HTML, EPUB）

在集成中，Docling 主要用来增强表格识别质量。
"""

from pathlib import Path
from typing import Optional


class DoclingEngine:
    """Docling 引擎 —— 表格增强 + 备选布局方案"""

    def __init__(self):
        self._converter = None

    def _get_converter(self):
        if self._converter is None:
            from docling.document_converter import DocumentConverter
            self._converter = DocumentConverter()
        return self._converter

    def convert(self, pdf_path: str) -> dict:
        """转换 PDF 为 Docling 文档结构

        Returns:
            {
                "doc": DoclingDocument (原始文档对象),
                "markdown": str (完整 MD),
                "tables": [{"cells": ..., "bbox": ..., "html": ...}],
                "blocks": [{"type": ..., "text": ..., "bbox": ...}],
            }
        """
        converter = self._get_converter()
        result = converter.convert(pdf_path)
        doc = result.document

        # 导出 Markdown
        markdown = doc.export_to_markdown()

        # 提取表格（TableFormer 的结果）
        tables = []
        for table in doc.tables:
            try:
                table_data = {
                    "label": table.label if hasattr(table, 'label') else "table",
                    "bbox": _bbox_to_dict(table.bbox) if hasattr(table, 'bbox') else None,
                    "data": table.export_to_dataframe(doc=doc) if hasattr(table, 'export_to_dataframe') else None,
                    "html": table.export_to_html(doc=doc) if hasattr(table, 'export_to_html') else None,
                }
                tables.append(table_data)
            except Exception:
                # 降级：手动提取表格文本
                text = ""
                try:
                    if hasattr(table, 'export_to_markdown'):
                        text = table.export_to_markdown()
                except Exception:
                    pass
                tables.append({"label": "table", "text": text})

        # 提取 block 列表
        blocks = []
        for item in doc.iterate_items():
            # Docling v2 迭代返回 (level, item) 元组
            if isinstance(item, tuple):
                level, item = item
            btype = item.label if hasattr(item, 'label') else "unknown"
            text = ""
            if hasattr(item, 'text'):
                text = item.text or ""
            elif hasattr(item, 'export_to_markdown'):
                try:
                    text = item.export_to_markdown()
                except Exception:
                    pass

            bbox = None
            if hasattr(item, 'prov') and item.prov:
                try:
                    bbox = _bbox_to_dict(item.prov[0].bbox) if hasattr(item.prov[0], 'bbox') else None
                except Exception:
                    pass

            blocks.append({
                "type": btype,
                "text": text,
                "bbox": bbox,
            })

        return {
            "doc": doc,
            "markdown": markdown,
            "tables": tables,
            "blocks": blocks,
            "metadata": {
                "engine": "Docling",
                "layout_model": "DocLayNet (heron-101)",
                "table_model": "TableFormer",
            },
        }

    def extract_tables(self, pdf_path: str) -> list[dict]:
        """专门提取表格 —— 发挥 TableFormer 最强表格识别能力"""
        result = self.convert(pdf_path)
        return result["tables"]


def _bbox_to_dict(bbox) -> dict | None:
    """BBox 对象 → dict"""
    if bbox is None:
        return None
    if isinstance(bbox, dict):
        return bbox
    try:
        return {"l": bbox.l, "t": bbox.t, "r": bbox.r, "b": bbox.b}
    except Exception:
        return None
