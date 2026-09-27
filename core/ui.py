# -*- coding: utf-8 -*-
"""终端交互 UI 辅助模块。

提供统一的标题、菜单、输入、提示输出封装，供各工具复用。
仅依赖标准库，避免因环境缺少第三方库而崩溃。

输出一律为纯文本，不使用 ANSI 颜色转义序列（因此不需要 `C` 颜色常量，
也不会在不同终端下出现乱码或多余的 ^[[0m 字符）。
Windows / Android 的差异只在 clear_screen() 的清屏命令上，由 core.runtime 决定。
"""

import os

from core import runtime

# ---------------------------------------------------------------- 基础输出

def clear_screen():
    """清屏（命令随运行平台变化：Windows 用 cls，Android 用 clear）。"""
    os.system(runtime.clear_command())


def line(char="-", width=56):
    print(char * width)


def title(text, width=56):
    """打印一级标题。"""
    print()
    print("=" * width)
    print(text.center(width - 2))
    print("=" * width)


def subtitle(text):
    """打印二级标题。"""
    print()
    print(text)
    line()


def info(msg):
    print(f"[信息] {msg}")


def success(msg):
    print(f"[成功] {msg}")


def warn(msg):
    print(f"[警告] {msg}")


def error(msg):
    print(f"[错误] {msg}")


# ---------------------------------------------------------------- 输入封装

def ask_raw(prompt):
    """读取一行输入，用于需要区分「空输入」与「输入流已结束」的场景。

    返回用户输入的字符串（可能为空串）；当输入流结束（EOF，例如
    `python main.py < NUL`、双击运行、管道输入用尽）或用户按下 Ctrl+C 时，
    返回 None。调用方必须处理 None，否则会陷入死循环。
    """
    try:
        return input(prompt + ": ").strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return None


def ask(prompt, default=None):
    """文本输入。default 为 None 时允许直接回车返回空字符串。

    输入流结束（EOF）或 Ctrl+C 时返回 default；default 为 None 则返回空串。
    注意：本函数无法区分「EOF」与「用户直接回车」，需要区分的场合请用 ask_raw()。
    """
    tip = f"{prompt}"
    if default is not None:
        tip += f"（默认 {default}）"
    value = ask_raw(tip)
    if value is None:                      # EOF / Ctrl+C
        return default if default is not None else ""
    if value == "" and default is not None:
        return default
    return value


def ask_int(prompt, default=None, min_value=None, max_value=None):
    """整数输入，非法时返回 default（若 default 为 None 则返回 None）。"""
    while True:
        raw = ask(prompt, default if default is not None else None)
        if raw == "" and default is None:
            return None
        try:
            value = int(raw)
        except ValueError:
            error("请输入一个整数。")
            continue
        if min_value is not None and value < min_value:
            error(f"数值不能小于 {min_value}。")
            continue
        if max_value is not None and value > max_value:
            error(f"数值不能大于 {max_value}。")
            continue
        return value


def ask_yes_no(prompt, default="n"):
    """是/否确认。"""
    while True:
        raw = ask(prompt + " (y/n)", default).lower()
        if raw in ("y", "yes", "是"):
            return True
        if raw in ("n", "no", "否"):
            return False
        error("请输入 y 或 n。")


def pause(msg="按回车键返回..."):
    try:
        input(msg)
    except (EOFError, KeyboardInterrupt):
        print()


# ---------------------------------------------------------------- 菜单

def menu(items, title_text="请选择功能", allow_back=True, back_label="返回上一级"):
    """打印并处理一个菜单。

    items: [(key, label, description?), ...]
           也可以是 (key, label) 二元组。
    返回: 用户选择的 key（字符串），或 None 表示返回上一级。
         输入流结束（EOF）或 Ctrl+C 时同样返回 None，保证不会死循环。
    """
    while True:
        print()
        print(title_text)
        line()
        for item in items:
            key, label = item[0], item[1]
            desc = item[2] if len(item) > 2 else ""
            if desc:
                print(f"  [{key}] {label}  {desc}")
            else:
                print(f"  [{key}] {label}")
        if allow_back:
            print(f"  [0] {back_label}")
        line()

        # 用 ask_raw 而非 ask：ask 会把 EOF 变成空串，空串不是合法序号，
        # while True 就会无限打印「无效的序号」（stdin 为 NUL / 管道用尽时）。
        choice = ask_raw("请输入序号")
        if choice is None:
            return None
        if choice == "0" and allow_back:
            return None
        valid = {str(i[0]) for i in items}
        if choice in valid:
            return choice
        error("无效的序号，请重新输入。")
