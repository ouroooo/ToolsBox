# -*- coding: utf-8 -*-
"""工具注册表。

每个工具模块需暴露：
    NAME      : 工具名（菜单中显示）
    DESCRIPTION: 简述（可选）
    ORDER     : 排序（可选，数字越小越前，默认 99）
    run()     : 工具入口函数，自行负责交互，返回后回到主菜单
"""

from tools import excel_extract
from tools import excel_multi_extract
from tools import image_to_pdf

# 全部工具平铺登记。新增工具时在此登记即可。
TOOLS = [
    excel_extract,
    image_to_pdf,
    excel_multi_extract,
]


def iter_tools():
    """返回全部工具（已按 ORDER 排序）。"""
    return sorted(TOOLS, key=lambda t: getattr(t, "ORDER", 99))


def find_tool(index):
    """按序号（从 1 开始）取工具模块。"""
    tools = iter_tools()
    if 1 <= index <= len(tools):
        return tools[index - 1]
    return None
