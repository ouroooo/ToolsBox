# -*- coding: utf-8 -*-
"""路径工具：项目根目录、输出目录，以及平台相关的常见目录与路径示例。

Android 与 Windows 的路径差异集中在本模块，通过 core.runtime 的当前平台判断：

    - common_dirs()  : Android 模式下提供 /sdcard 等常见目录；Windows 模式返回空
    - example_path() : 交互提示里用的示例路径，随平台变化
"""

import os

from core import runtime

# 项目根目录（core/ 的上一级）
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Android / Termux 终端常见的外部存储目录，供「浏览常见目录」快速选择。
# Windows 下不存在这些路径，因此只在 Android 模式下启用（见 common_dirs()）。
ANDROID_COMMON_DIRS = [
    "/sdcard/Download",
    "/sdcard/Documents",
    "/storage/emulated/0/Download",
    "/storage/emulated/0/Documents",
    "/sdcard",
    "/storage/emulated/0",
]


def common_dirs():
    """返回当前平台可快速浏览的常见目录；Windows 模式返回空列表。"""
    return list(ANDROID_COMMON_DIRS) if runtime.is_android() else []


def has_common_dirs():
    """当前平台是否提供「浏览常见目录」功能（仅 Android 模式为真）。"""
    return runtime.is_android()


def example_path(filename="成绩表.xlsx"):
    """给出当前平台下直观的路径示例，用于交互提示。"""
    if runtime.is_windows():
        return "C:\\Users\\Public\\Documents\\" + filename
    return "/sdcard/Download/" + filename


def ensure_output_dir():
    """确保输出目录存在。"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    return OUTPUT_DIR


def default_output_name(source_path, suffix):
    """根据源文件名生成默认输出文件名，例如 班级_筛选.xlsx。"""
    stem = os.path.splitext(os.path.basename(source_path))[0]
    ext = os.path.splitext(source_path)[1] or ".xlsx"
    return f"{stem}_{suffix}{ext}"


def normalize_path(raw):
    """清理用户输入路径：去引号、去空白、展开 ~。"""
    if raw is None:
        return ""
    p = raw.strip().strip("'").strip('"')
    p = os.path.expanduser(p)
    return p


def is_readable_file(path):
    return bool(path) and os.path.isfile(path) and os.access(path, os.R_OK)
