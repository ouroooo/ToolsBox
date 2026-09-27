# -*- coding: utf-8 -*-
"""图片转 PDF 核心逻辑。

流程：
    1. 收集待转换的图片（逐个添加 / 按目录批量添加 / 粘贴多行路径）
    2. 调整顺序、删除误加入的项
    3. 选择输出尺寸模式（原图大小 / A4 自适应 / 固定宽度）
    4. 预览并导出 PDF 到 output/ 目录

支持 jpg / jpeg / png（另兼容 bmp / gif / webp / tiff，取决于 Pillow 解码能力）。
依赖 Pillow（PIL）。缺失时给出友好提示，不影响工具箱其他功能。
"""

import os

from core import ui, paths

# ---------------------------------------------------------------- 常量

SUPPORTED_EXTS = (
    ".jpg", ".jpeg", ".png",       # 需求明确的三种
    ".bmp", ".gif", ".webp", ".tif", ".tiff",  # 顺带兼容
)

PAGE_MODES = [
    ("1", "原图尺寸", "每页与图片像素一致（1px = 1pt），不缩放"),
    ("2", "A4 自适应", "按 A4 纵向排版，图片等比缩放并居中留白"),
    ("3", "固定宽度", "指定页面宽度(mm)，高度按图片比例计算"),
]

A4_WIDTH_MM = 210.0
A4_HEIGHT_MM = 240.0  # 留出上下边距后的可用高度
A4_MARGIN_MM = 15.0


# ---------------------------------------------------------------- 依赖检查

def _load_pillow():
    """延迟导入 Pillow，缺失时给出安装指引。"""
    try:
        from PIL import Image
        return Image
    except ImportError:
        ui.error("未检测到 Pillow 库，无法处理图片。")
        ui.info("请先安装：pip install Pillow")
        return None


# ---------------------------------------------------------------- 文件收集

def _list_dir_images(directory, recursive=False):
    """列出目录下的图片文件，返回绝对路径列表。"""
    found = []
    try:
        for entry in sorted(os.listdir(directory)):
            full = os.path.join(directory, entry)
            if os.path.isfile(full) and entry.lower().endswith(SUPPORTED_EXTS):
                if not entry.startswith("~$"):
                    found.append(full)
            elif recursive and os.path.isdir(full):
                found.extend(_list_dir_images(full, recursive=True))
    except OSError:
        pass
    return found


def _is_image(path):
    return path.lower().endswith(SUPPORTED_EXTS)


def _add_single_file(items):
    """添加一张图片：输入完整路径；Android 模式下可先浏览常见目录。"""
    while True:
        # 「浏览常见目录」只在 Android 模式提供：Windows 下不存在 /sdcard 等目录
        if paths.has_common_dirs() and ui.ask_yes_no("是否浏览常见目录挑选", "n"):
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
            files = _list_dir_images(directory)
            if not files:
                ui.warn(f"{directory} 中没有图片文件。")
                continue
            fpick = ui.menu(
                [(str(i + 1), os.path.basename(f)) for i, f in enumerate(files)],
                title_text=f"{directory} 中的图片",
                allow_back=True,
            )
            if fpick is None:
                continue
            path = files[int(fpick) - 1]
            if path in items:
                ui.warn("该图片已在列表中。")
                return False
            items.append(path)
            ui.success(f"已添加：{os.path.basename(path)}")
            return True

        raw = ui.ask("图片路径（直接回车取消）")
        p = paths.normalize_path(raw)
        if not p:
            return False
        if not os.path.isfile(p):
            ui.error("文件不存在，请重新输入。")
            continue
        if not _is_image(p):
            ui.warn("该文件后缀不在支持范围内。")
            if not ui.ask_yes_no("仍要添加吗", "n"):
                continue
        if p in items:
            ui.warn("该图片已在列表中。")
            return False
        items.append(p)
        ui.success(f"已添加：{os.path.basename(p)}")
        return True


def _add_many_files(items):
    """一次粘贴多行路径（也支持空格/逗号分隔）。"""
    print("  每行一个路径，也支持空格或逗号分隔；输入空行结束。")
    added = 0
    while True:
        raw = ui.ask("路径（空行结束）")
        if not raw:
            break
        for token in raw.replace(",", " ").split():
            p = paths.normalize_path(token)
            if not p:
                continue
            if not os.path.isfile(p):
                ui.error(f"跳过（文件不存在）：{p}")
                continue
            if not _is_image(p):
                ui.warn(f"跳过（不是图片）：{p}")
                continue
            if p in items:
                ui.warn(f"跳过（已存在）：{os.path.basename(p)}")
                continue
            items.append(p)
            added += 1
    if added:
        ui.success(f"已批量添加 {added} 张图片。")
    else:
        ui.warn("没有添加任何图片。")
    return added > 0


def _add_directory(items):
    """整个目录批量加入（可选递归子目录）。"""
    raw = ui.ask("目录路径")
    directory = paths.normalize_path(raw)
    if not directory or not os.path.isdir(directory):
        ui.error("目录不存在。")
        return False
    recursive = ui.ask_yes_no("是否包含子目录", "n")
    files = _list_dir_images(directory, recursive=recursive)
    if not files:
        ui.warn("该目录下没有图片文件。")
        return False

    if ui.ask_yes_no(f"共找到 {len(files)} 张图片，全部添加", "y"):
        chosen = list(range(len(files)))
    else:
        print("  输入要添加的序号，如 1-5,8；直接回车取消。")
        spec = ui.ask("序号范围")
        chosen = _parse_index_spec(spec, len(files))
        if chosen is None:
            ui.warn("已取消。")
            return False

    added = 0
    for i in chosen:
        p = files[i]
        if p in items:
            continue
        items.append(p)
        added += 1
    if added:
        ui.success(f"已添加 {added} 张图片（跳过 {len(files) - added} 张重复项）。")
    else:
        ui.warn("全部图片都已在列表中。")
    return added > 0


def _parse_index_spec(spec, total):
    """解析 "1-5,8" 形式的序号，返回 0 基索引列表；非法返回 None。"""
    spec = (spec or "").strip()
    if not spec:
        return None
    result = []
    for chunk in spec.replace("，", ",").split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            parts = chunk.split("-", 1)
            try:
                start, end = int(parts[0]), int(parts[1])
            except ValueError:
                return None
            if start > end:
                start, end = end, start
            for n in range(start, end + 1):
                if 1 <= n <= total:
                    result.append(n - 1)
        else:
            try:
                n = int(chunk)
            except ValueError:
                return None
            if 1 <= n <= total:
                result.append(n - 1)
    if not result:
        return None
    # 去重且保持顺序
    seen, ordered = set(), []
    for i in result:
        if i not in seen:
            seen.add(i)
            ordered.append(i)
    return ordered


def _collect_images():
    """主收集流程。返回图片路径列表（可能为空）。"""
    ui.subtitle("步骤 1 / 4：添加图片")
    print(f"  支持格式：{', '.join(SUPPORTED_EXTS)}")

    items = []
    while True:
        if items:
            _print_items(items)
        else:
            ui.info("当前列表为空。")

        pick = ui.menu(
            [
                ("1", "添加单张图片",
                 "输入路径或浏览目录挑选" if paths.has_common_dirs() else "输入完整路径"),
                ("2", "批量粘贴路径", "一次输入多个路径"),
                ("3", "添加整个目录", "可选是否包含子目录"),
                ("4", "调整顺序", "上移 / 下移某张图片"),
                ("5", "删除某项", "从列表中移除"),
                ("6", "清空列表", "移除全部图片"),
                ("7", "完成添加，进入下一步", f"当前 {len(items)} 张"),
            ],
            title_text="图片列表操作",
            allow_back=True,
            back_label="取消并返回",
        )
        if pick is None:
            if items and not ui.ask_yes_no("列表未保存，确定放弃吗", "n"):
                continue
            return []

        if pick == "1":
            _add_single_file(items)
        elif pick == "2":
            _add_many_files(items)
        elif pick == "3":
            _add_directory(items)
        elif pick == "4":
            _move_item(items)
        elif pick == "5":
            _remove_item(items)
        elif pick == "6":
            if items and ui.ask_yes_no("确定清空列表", "n"):
                items.clear()
                ui.success("列表已清空。")
        elif pick == "7":
            if not items:
                ui.error("请先添加至少一张图片。")
                continue
            return items


def _print_items(items):
    """打印当前图片列表。"""
    print()
    print(f"当前图片列表（共 {len(items)} 张）")
    ui.line()
    for i, p in enumerate(items, start=1):
        print(f"  [{i}] {os.path.basename(p)}  {p}")
    ui.line()


def _move_item(items):
    """调整某张图片的位置。"""
    if len(items) < 2:
        ui.warn("至少需要两张图片才能调整顺序。")
        return
    _print_items(items)
    idx = ui.ask_int("要移动的序号（回车取消）", min_value=1, max_value=len(items))
    if idx is None:
        return
    target = ui.ask_int(f"移动到第几位（1-{len(items)}）",
                        min_value=1, max_value=len(items))
    if target is None:
        return
    item = items.pop(idx - 1)
    items.insert(target - 1, item)
    ui.success(f"已将 {os.path.basename(item)} 移动到第 {target} 位。")


def _remove_item(items):
    """从列表中删除一项。"""
    if not items:
        ui.warn("列表为空。")
        return
    _print_items(items)
    idx = ui.ask_int("要删除的序号（回车取消）", min_value=1, max_value=len(items))
    if idx is None:
        return
    removed = items.pop(idx - 1)
    ui.success(f"已移除：{os.path.basename(removed)}")


# ---------------------------------------------------------------- 输出选项

def _ask_page_mode():
    """选择输出尺寸模式，返回 (mode_key, width_mm)。"""
    ui.subtitle("步骤 2 / 4：选择输出尺寸")
    pick = ui.menu(
        PAGE_MODES,
        title_text="尺寸模式",
        allow_back=False,
    )
    if pick == "3":
        width = ui.ask_int("页面宽度(mm)", default=210, min_value=20, max_value=2000)
        return pick, float(width or 210)
    return pick, None


def _ask_output_path(first_image):
    """确定输出文件路径。"""
    ui.subtitle("步骤 3 / 4：选择输出位置")
    paths.ensure_output_dir()
    default_name = os.path.splitext(os.path.basename(first_image))[0] + "_合并.pdf"
    default_path = os.path.join(paths.OUTPUT_DIR, default_name)
    ui.info(f"默认输出：{default_path}")

    if ui.ask_yes_no("使用默认输出路径", "y"):
        if os.path.exists(default_path):
            if not ui.ask_yes_no(f"{default_name} 已存在，是否覆盖", "n"):
                stem, ext = os.path.splitext(default_path)
                n = 2
                while os.path.exists(f"{stem}_{n}{ext}"):
                    n += 1
                default_path = f"{stem}_{n}{ext}"
                ui.info(f"将保存为：{os.path.basename(default_path)}")
        return default_path

    while True:
        raw = ui.ask("请输入输出文件路径（.pdf）")
        p = paths.normalize_path(raw)
        if not p:
            ui.warn("路径为空，改回默认输出。")
            return default_path
        if not p.lower().endswith(".pdf"):
            p += ".pdf"
        out_dir = os.path.dirname(os.path.abspath(p))
        if out_dir and not os.path.isdir(out_dir):
            if ui.ask_yes_no(f"目录不存在：{out_dir}，是否创建", "y"):
                try:
                    os.makedirs(out_dir, exist_ok=True)
                except OSError as exc:
                    ui.error(f"创建目录失败：{exc}")
                    continue
            else:
                continue
        if os.path.exists(p):
            if not ui.ask_yes_no("文件已存在，是否覆盖", "n"):
                continue
        return p


# ---------------------------------------------------------------- 转换

def _open_rgb(Image, path):
    """打开图片并统一为 RGB / L 模式（JPEG 不支持透明通道）。"""
    img = Image.open(path)
    if img.mode in ("RGBA", "LA", "P"):
        # 透明部分以白底填充，避免转 PDF 时出现黑块
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        return background
    if img.mode not in ("RGB", "L"):
        return img.convert("RGB")
    return img


def _convert(items, mode, width_mm, out_path):
    """执行转换。返回 (页数, 输出路径)；失败抛异常。"""
    Image = _load_pillow()
    if Image is None:
        return None

    mm2pt = 72.0 / 25.4
    frames = []

    ui.info("正在读取图片...")
    for path in items:
        try:
            img = _open_rgb(Image, path)
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"无法读取图片 {os.path.basename(path)}：{exc}")

        if mode == "1":
            # 原图尺寸：1px = 1pt
            frames.append(img)
        elif mode == "2":
            # A4 自适应：等比缩放到可用区域内居中
            avail_w = (A4_WIDTH_MM - A4_MARGIN_MM * 2) * mm2pt
            avail_h = (A4_HEIGHT_MM - A4_MARGIN_MM * 2) * mm2pt
            ratio = min(avail_w / img.width, avail_h / img.height)
            new_size = (max(1, int(img.width * ratio)),
                        max(1, int(img.height * ratio)))
            canvas = Image.new("RGB", (int(A4_WIDTH_MM * mm2pt),
                                       int(A4_HEIGHT_MM * mm2pt)),
                               (255, 255, 255))
            resized = img.resize(new_size, Image.LANCZOS)
            offset = ((canvas.width - new_size[0]) // 2,
                      (canvas.height - new_size[1]) // 2)
            canvas.paste(resized, offset)
            frames.append(canvas)
        else:
            # 固定宽度：按指定宽度(mm)等比缩放
            new_w = max(1, int(width_mm * mm2pt))
            ratio = new_w / img.width
            new_h = max(1, int(img.height * ratio))
            frames.append(img.resize((new_w, new_h), Image.LANCZOS))

    if not frames:
        raise RuntimeError("没有可转换的图片。")

    first, rest = frames[0], frames[1:]
    ui.info(f"正在生成 PDF（共 {len(frames)} 页）...")
    first.save(out_path, "PDF", resolution=150.0, save_all=bool(rest),
               append_images=rest)
    return len(frames), out_path


# ---------------------------------------------------------------- 入口

def run():
    """工具入口。"""
    ui.title("图片转 PDF")

    Image = _load_pillow()
    if Image is None:
        ui.pause("按回车继续...")
        return

    # 1. 收集图片
    items = _collect_images()
    if not items:
        ui.warn("已取消，未生成文件。")
        ui.pause("按回车继续...")
        return

    # 2. 尺寸模式
    mode, width_mm = _ask_page_mode()
    if mode is None:                       # 输入流结束（EOF）/ Ctrl+C
        ui.warn("已取消，未生成文件。")
        ui.pause("按回车继续...")
        return

    # 3. 输出路径
    out_path = _ask_output_path(items[0])

    # 4. 预览
    ui.subtitle("步骤 4 / 4：确认并导出")
    mode_label = next((m[1] for m in PAGE_MODES if m[0] == mode), mode)
    ui.info(f"图片数量：{len(items)} 张")
    ui.info(f"尺寸模式：{mode_label}" + (f"（宽度 {width_mm:.0f} mm）" if width_mm else ""))
    ui.info(f"输出文件：{out_path}")

    total_size = 0
    for p in items:
        try:
            total_size += os.path.getsize(p)
        except OSError:
            pass
    ui.info(f"源文件总大小：{total_size / 1024 / 1024:.2f} MB")

    if not ui.ask_yes_no("确认开始转换", "y"):
        ui.warn("已取消。")
        ui.pause("按回车继续...")
        return

    try:
        result = _convert(items, mode, width_mm, out_path)
    except Exception as exc:  # noqa: BLE001
        ui.error(f"转换失败：{exc}")
        ui.pause("按回车继续...")
        return

    if not result:
        ui.pause("按回车继续...")
        return

    pages, real_path = result
    size_mb = os.path.getsize(real_path) / 1024 / 1024
    ui.success(f"已生成：{real_path}")
    ui.info(f"共 {pages} 页，文件大小 {size_mb:.2f} MB。")

    if ui.ask_yes_no("是否继续转换其他图片", "n"):
        run()
    else:
        ui.pause("按回车继续...")