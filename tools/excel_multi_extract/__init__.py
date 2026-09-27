# -*- coding: utf-8 -*-
"""多文件 / 多工作表行列提取工具包。

对外暴露本工具的元信息与入口 run()。
实现细节见 extractor.py。
"""

from .extractor import run

NAME = "多表行列提取"
ORDER = 3
DESCRIPTION = "多 xlsx / 多工作表按行或列条件提取并合并"

__all__ = ["run", "NAME", "ORDER", "DESCRIPTION"]