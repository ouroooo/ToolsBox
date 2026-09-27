#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""工具箱主程序（Windows / Android 双平台）。

启动后先询问运行平台，再展示全部工具列表，用户输入序号即可执行对应工具。

运行（两个平台命令相同）：
    python main.py

如需跳过平台询问（自动化脚本 / 快捷方式），可显式指定：
    python main.py --platform=windows
    python main.py --platform=android

说明：第 1 行的 shebang 只对 Unix / Android 终端生效，Windows cmd 不解析它，
保留不影响运行；Windows 下用 python main.py 启动即可。
"""

import os
import sys

# 确保可以从任意工作目录启动
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import ui            # noqa: E402
from core import registry      # noqa: E402
from core import runtime       # noqa: E402

APP_NAME = "ToolsBox"
VERSION = "0.3.0"

PLATFORM_FLAG = "--platform="


# ---------------------------------------------------------------- 欢迎界面

def show_banner():
    ui.title(APP_NAME)
    print(f"  版本 {VERSION}  ·  平台 {runtime.label()}"
          f"  ·  输入序号选择工具  ·  输入 0 退出")


# ---------------------------------------------------------------- 主菜单

def main_menu():
    """主菜单：平铺展示全部工具，直接选择执行。"""
    tools = registry.iter_tools()
    if not tools:
        ui.warn("当前还没有注册任何工具。")
        return

    items = []
    for i, t in enumerate(tools):
        items.append((str(i + 1), getattr(t, "NAME", t.__name__),
                      getattr(t, "DESCRIPTION", "")))

    while True:
        show_banner()
        pick = ui.menu(
            items,
            title_text="请选择工具",
            allow_back=True,
            back_label="退出",
        )
        if pick is None:
            return

        tool = tools[int(pick) - 1]
        name = getattr(tool, "NAME", tool.__name__)
        run_tool(tool, name)


def run_tool(tool, name):
    """执行工具，统一捕获异常，保证不会直接退出主程序。"""
    print()
    ui.subtitle(f"正在运行：{name}")
    try:
        tool.run()
    except KeyboardInterrupt:
        print()
        ui.warn("操作已被用户中断。")
        ui.pause()
    except Exception as exc:  # noqa: BLE001
        ui.error(f"工具执行出错：{exc}")
        import traceback
        print(traceback.format_exc())
        ui.pause()


# ---------------------------------------------------------------- 平台选择

def platform_from_argv(argv):
    """从命令行参数里取出 --platform= 的值；未指定时返回 None。"""
    for arg in argv:
        if arg.startswith(PLATFORM_FLAG):
            return arg[len(PLATFORM_FLAG):]
    return None


def decide_platform(argv):
    """决定本次运行的平台。

    - 命令行给了 --platform=windows|android 且取值合法：直接采用，不询问；
    - 取值非法：提示后回退到交互询问；
    - 未指定：交互询问（默认值来自自动探测，直接回车即采用）。
    """
    requested = platform_from_argv(argv)
    if requested is None:
        return runtime.ask_platform()

    picked = runtime.parse_platform_arg(requested)
    if picked is None:
        ui.error(f"无法识别的平台：{requested}（可用：windows / android）")
        return runtime.ask_platform()

    runtime.set_platform(picked)
    ui.info(f"运行平台：{runtime.label()}（由命令行 --platform 指定）")
    return picked


# ---------------------------------------------------------------- 入口

def main():
    decide_platform(sys.argv[1:])
    print()

    try:
        main_menu()
    except (KeyboardInterrupt, EOFError):
        print()
    ui.info("感谢使用，再见！")


if __name__ == "__main__":
    main()
