# -*- coding: utf-8 -*-
"""图片转 PDF 工具包。

对外暴露本工具的元信息与入口 run()。
实现细节见 converter.py。
"""

from .converter import run

NAME = "图片转 PDF"
ORDER = 2
DESCRIPTION = "单个或多个 jpg/png/jpeg 合并为 PDF"

__all__ = ["run", "NAME", "ORDER", "DESCRIPTION"]
