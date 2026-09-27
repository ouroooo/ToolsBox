# -*- coding: utf-8 -*-
"""Excel 数据提取工具包。

对外暴露本工具的元信息与入口 run()。
实现细节见 extractor.py。
"""

from .extractor import run

NAME = "xlsx 数据提取"
ORDER = 1
DESCRIPTION = "按列筛选（如班级=15）后导出新表"

__all__ = ["run", "NAME", "ORDER", "DESCRIPTION"]