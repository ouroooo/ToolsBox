# -*- coding: utf-8 -*-
"""PDF 单页提取核心逻辑。

流程：
    1. 选择 PDF 文件（输入完整路径；Android 模式下还可浏览常见目录）
    2. 读取并显示该 PDF 的总页数
    3. 指定要提取的页码（1 起，带范围校验）
    4. 确认输出路径（默认输出到 output/ 目录）
    5. 另存为只含该页的新 PDF

依赖 pypdf。缺失时给出安装指引，不影响工具箱其它功能。
"""

import os

from core import ui, paths

# ---------------------------------------------------------------- 依赖检查

def _load_pypdf():
    """延迟导入 pypdf，缺失时给出安装指引。"""
    try:
        from pypdf import PdfReader, PdfWriter
        return PdfReader, PdfWriter
    except ImportError:
        ui.error("未检测到 pypdf 库，无法处理 PDF 文件。")
        ui.info("请先安装：pip install pypdf")
        return None


# ---------------------------------------------------------------- 文件选择

def _list_dir_pdfs(directory):
    """列出目录下的 PDF 文件。"""
    try:
        return sorted(
            f for f in os.listdir(directory)
            if f.lower().endswith(".pdf") and not f.startswith("~$")
        )
    except OSError:
        return []


def _choose_pdf():
    """让用户输入 PDF 路径；Android 模式下可先浏览常见目录。

    返回绝对路径或 None（取消）。
    """
    ui.subtitle("步骤 1 / 4：选择 PDF 文件")
    # 「浏览常见目录」只在 Android 模式提供：Windows 下不存在 /sdcard 等目录
    if paths.has_common_dirs():
        print("  直接输入完整路径，或先浏览常见目录。")
    else:
        print("  请输入 PDF 文件的完整路径（路径含空格时可用引号包住）。")
    print(f"  示例：{paths.example_path('成绩单.pdf')}")

    while True:
        if paths.has_common_dirs() and ui.ask_yes_no("是否浏览常见目录", "n"):
            dirs = paths.common_dirs()
            pick = ui.menu(
                [(str(i + 1), d) for i, d in enumerate(dirs)],
                title_text="选择目录",
                allow_back=True,
            )
            if pick is None:
                continue
            directory = dirs[int(pick) - 1]
            if not os.path.isdir(directory):
                ui.error(f"目录不存在：{directory}")
                continue
            files = _list_dir_pdfs(directory)
            if not files:
                ui.warn(f"{directory} 中没有 PDF 文件。")
                continue
            fpick = ui.menu(
                [(str(i + 1), f) for i, f in enumerate(files)],
                title_text=f"{directory} 中的 PDF",
                allow_back=True,
            )
            if fpick is None:
                continue
            return os.path.join(directory, files[int(fpick) - 1])

        raw = ui.ask("文件路径（直接回车取消）")
        p = paths.normalize_path(raw)
        if not p:
            ui.warn("已取消。")
            return None
        if not os.path.isfile(p):
            ui.error("文件不存在，请重新输入。")
            continue
        if not paths.is_readable_file(p):
            ui.error("文件不可读，请检查权限。")
            continue
        if not p.lower().endswith(".pdf"):
            ui.warn("该文件不是 .pdf，可能无法解析。")
            if not ui.ask_yes_no("仍要继续吗", "n"):
                continue
        return p


# ---------------------------------------------------------------- 页码与输出

def _ask_page_num(total_pages):
    """询问要提取的页码（1 起）。返回页码，或 None（取消 / 输入流结束）。"""
    ui.subtitle("步骤 3 / 4：选择页码")
    page_num = ui.ask_int("要提取的页码", min_value=1, max_value=total_pages)
    if page_num is None:
        ui.warn("已取消。")
        return None
    return page_num


def _ask_output_path(input_path, page_num):
    """确定输出文件路径。返回路径，或 None（取消）。"""
    ui.subtitle("步骤 4 / 4：确认输出")
    default_name = paths.default_output_name(input_path, f"page{page_num}")
    default_path = os.path.join(paths.ensure_output_dir(), default_name)
    ui.info(f"默认输出：{default_path}")

    out_path = ui.ask("输出文件路径", default=default_path)
    out_path = paths.normalize_path(out_path)
    if not out_path:
        ui.warn("输出路径为空，已取消。")
        return None
    if not out_path.lower().endswith(".pdf"):
        out_path += ".pdf"

    out_dir = os.path.dirname(os.path.abspath(out_path))
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError as exc:
        ui.error(f"无法创建输出目录：{exc}")
        return None
    return out_path


# ---------------------------------------------------------------- 提取

def _extract_page(PdfReader, PdfWriter, input_path, page_num, output_path):
    """把 input_path 的第 page_num 页写入 output_path。"""
    reader = PdfReader(input_path)
    writer = PdfWriter()
    writer.add_page(reader.pages[page_num - 1])
    with open(output_path, "wb") as f:
        writer.write(f)


# ---------------------------------------------------------------- 入口

def run():
    """工具入口。"""
    ui.title("PDF 单页提取")

    loaded = _load_pypdf()
    if loaded is None:
        ui.pause()
        return
    PdfReader, PdfWriter = loaded

    # 1. 选文件
    input_path = _choose_pdf()
    if not input_path:
        return

    # 2. 读取总页数
    ui.subtitle("步骤 2 / 4：读取 PDF")
    try:
        reader = PdfReader(input_path)
        total_pages = len(reader.pages)
    except Exception as exc:  # noqa: BLE001
        ui.error(f"无法读取 PDF：{exc}")
        ui.pause()
        return

    if total_pages <= 0:
        ui.error("该 PDF 没有任何页面，无法提取。")
        ui.pause()
        return

    ui.info(f"文件：{input_path}")
    ui.info(f"共 {total_pages} 页。")

    # 3. 页码
    page_num = _ask_page_num(total_pages)
    if page_num is None:
        return

    # 4. 输出路径
    out_path = _ask_output_path(input_path, page_num)
    if not out_path:
        return

    # 执行提取
    try:
        _extract_page(PdfReader, PdfWriter, input_path, page_num, out_path)
    except PermissionError:
        ui.error("写入被拒绝，目标文件可能正被占用，或路径无写权限。")
        ui.pause()
        return
    except Exception as exc:  # noqa: BLE001
        ui.error(f"提取失败：{exc}")
        ui.pause()
        return

    print()
    ui.success(f"已生成：{out_path}")
    ui.info(f"已提取第 {page_num} 页（原文件共 {total_pages} 页）。")

    if ui.ask_yes_no("是否继续提取其他页面", "n"):
        run()
        return
    ui.pause()
