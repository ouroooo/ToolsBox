# -*- coding: utf-8 -*-
"""PDF 单页提取工具包。

对外暴露本工具的元信息与入口 run()。
实现细节见 pdfpage.py。
"""

from .pdfpage import run

NAME = "PDF 单页提取"
ORDER = 4
DESCRIPTION = "把 PDF 中的指定一页另存为新的 PDF"

__all__ = ["run", "NAME", "ORDER", "DESCRIPTION"]
