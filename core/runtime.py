# -*- coding: utf-8 -*-
"""运行平台识别与切换。

同一套代码同时支持 Windows 与 Android(Termux)，两者的差异集中在三处：

    1. 清屏命令 —— Windows 用 cls，Android 用 clear（见 clear_command()）
    2. 常见目录 —— Android 可浏览 /sdcard、/storage/emulated/0 等目录，
                   Windows 无此概念（见 core/paths.py 的 common_dirs()）
    3. 路径示例 —— 交互提示里的示例路径形式不同（见 core/paths.py 的 example_path()）

启动时由 main.py 调用 ask_platform() 询问用户；默认值取自 detect() 的自动探测
结果，直接回车即采用。其它模块通过 is_windows() / is_android() 读取当前平台。

注意：本模块**不在顶层导入任何项目内模块**，以免与 core.ui 形成循环导入
（ui.clear_screen() 需要 clear_command()）。ask_platform() 要用到 ui，
因此在函数内部延迟导入。
"""

import os

WINDOWS = "windows"
ANDROID = "android"

# 当前生效的平台；None 表示尚未选择，此时按 detect() 的结果处理
_platform = None


# ---------------------------------------------------------------- 探测与设置

def detect():
    """自动探测当前平台。

    Windows 用 os.name 判断；其余环境（Android/Termux、Linux、macOS）统一归入
    android 分支 —— 它们都是 POSIX 环境，清屏命令同为 clear。
    """
    return WINDOWS if os.name == "nt" else ANDROID


def current():
    """返回当前生效的平台标识。"""
    return _platform if _platform is not None else detect()


def set_platform(name):
    """设置当前平台，name 需为 WINDOWS 或 ANDROID。"""
    global _platform
    if name not in (WINDOWS, ANDROID):
        raise ValueError(f"未知平台：{name!r}")
    _platform = name
    return _platform


def is_windows():
    """当前是否为 Windows 模式。"""
    return current() == WINDOWS


def is_android():
    """当前是否为 Android / Termux 模式。"""
    return current() == ANDROID


def label(name=None):
    """平台的中文显示名。"""
    return "Windows" if (name or current()) == WINDOWS else "Android / Termux"


def clear_command():
    """当前平台的清屏命令。"""
    return "cls" if is_windows() else "clear"


def parse_platform_arg(value):
    """把命令行传入的平台写法归一化；无法识别时返回 None。"""
    if value is None:
        return None
    v = str(value).strip().lower()
    if v in ("windows", "win", "w", "nt", "1"):
        return WINDOWS
    if v in ("android", "termux", "linux", "posix", "a", "2"):
        return ANDROID
    return None


# ---------------------------------------------------------------- 启动时询问

def ask_platform():
    """启动时询问用户运行平台，返回选中的平台标识。

    默认项来自 detect()：直接回车即采用自动探测结果。
    输入流已结束（EOF）或按 Ctrl+C 时同样采用默认值，不会卡住或死循环。
    """
    from core import ui            # 延迟导入，避免与 core.ui 循环导入

    default_key = "1" if detect() == WINDOWS else "2"

    ui.subtitle("运行平台")
    print("  本工具箱同时支持 Windows 与 Android，请确认当前运行平台。")
    print("  两种模式的差别：清屏命令、是否提供「浏览常见目录」、提示里的路径示例。")
    print()
    ui.line()
    print("    [1] Windows")
    print("        使用 cls 清屏；路径形如 C:\\Users\\Public\\Documents\\成绩表.xlsx")
    print("    [2] Android / Termux")
    print("        使用 clear 清屏；可浏览 /sdcard、/storage/emulated/0 等常见目录")
    ui.line()

    while True:
        raw = ui.ask("请输入序号", default=default_key)
        if raw in ("1", "2"):
            break
        ui.error("请输入 1 或 2。")

    set_platform(WINDOWS if raw == "1" else ANDROID)
    ui.success(f"当前运行平台：{label()}")
    return current()
