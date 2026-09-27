# -*- coding: utf-8 -*-
"""多文件 / 多工作表行列提取核心逻辑。

需求：
    1. 用户可导入多个 xlsx，或一个/多个 xlsx 中的多个工作表；
    2. 对每个工作表，按「行条件」和/或「列条件」提取数据（至少一个）；
       - 行条件：如「某列包含 15」的所有数据行；
       - 列条件：如「列名包含 15」的整列；
       - 行 + 列：同时满足，取交集。
    3. 把所有工作表的结果合并到一个新的 xlsx。

流程：
    1. 收集源文件（可多个）
    2. 为每个文件选择工作表（可多选 / 全部）
    3. 设置表头行范围（多行表头）
    4. 设置行条件（0 个或多个，多个为 AND）
    5. 设置列条件（0 个或多个，不设则保留全部列）
    6. 预览各表命中量
    7. 合并导出到一个新 xlsx（每个工作表一个 sheet，同名自动加后缀）

依赖 openpyxl。
"""

import os

from core import ui, paths

# ---------------------------------------------------------------- 依赖检查

def _load_openpyxl():
    """延迟导入 openpyxl，缺失时给出安装指引。"""
    try:
        import openpyxl
        return openpyxl
    except ImportError:
        ui.error("未检测到 openpyxl 库，无法处理 xlsx 文件。")
        ui.info("请先安装：pip install openpyxl")
        return None


# ---------------------------------------------------------------- 通用小工具

def _col_letter(idx):
    """1 -> A, 27 -> AA。"""
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _cell_text(value):
    """把单元格值转成用于显示/比较的字符串。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _merged_value_map(ws):
    """收集合并单元格的左上角值，返回 {(row, col): 值}。"""
    mapping = {}
    for rng in ws.merged_cells.ranges:
        anchor = ws.cell(row=rng.min_row, column=rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                mapping[(r, c)] = anchor
    return mapping


def _header_cell_text(ws, merged, row, col):
    """读取表头单元格文字，兼容合并单元格。"""
    if (row, col) in merged:
        return _cell_text(merged[(row, col)])
    return _cell_text(ws.cell(row=row, column=col).value)


def _build_column_name(parts):
    """把多行表头片段拼接成列名（相邻重复片段去重）。"""
    cleaned = []
    for p in parts:
        p = (p or "").strip()
        if not p:
            continue
        if cleaned and cleaned[-1] == p:
            continue
        cleaned.append(p)
    return " / ".join(cleaned)


def _read_header(ws, header_rows):
    """读取表头。返回 [(列序号, 列名), ...]，跳过整段表头全空的列。"""
    merged = _merged_value_map(ws)
    cols = []
    for c in range(1, ws.max_column + 1):
        parts = [_header_cell_text(ws, merged, r, c) for r in header_rows]
        name = _build_column_name(parts)
        if name:
            cols.append((c, name))
    return cols


def _parse_index_spec(raw, count):
    """解析 "1-5,8" 形式的序号选择（支持中文逗号）。

    返回去重且保持输入顺序的 1-based 序号列表；非法输入返回 None。
    count 为可选总数，用于边界校验。
    """
    raw = (raw or "").replace("，", ",").replace(" ", "")
    if raw == "":
        return []
    result = []
    for part in raw.split(","):
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            if not a.isdigit() or not b.isdigit():
                return None
            a, b = int(a), int(b)
            if a > b:
                a, b = b, a
            rng = range(a, b + 1)
        else:
            if not part.isdigit():
                return None
            rng = [int(part)]
        for n in rng:
            if n < 1 or (count is not None and n > count):
                return None
            if n not in result:
                result.append(n)
    return result


# ---------------------------------------------------------------- 匹配条件

MATCH_MODES = [
    ("1", "精确匹配", "单元格内容与目标值完全相等，如 15 == 15"),
    ("2", "包含关键词", "单元格内容包含目标值，如 15班 含 15"),
    ("3", "不等于", "单元格内容与目标值不相等"),
    ("4", "为空", "单元格为空（无目标值）"),
    ("5", "非空", "单元格有内容（无目标值）"),
]

_MODE_LABEL = {k: v for k, v, _ in MATCH_MODES}


def _match(text, mode, target):
    """执行单值匹配判定。"""
    if mode == "1":
        return text == target
    if mode == "2":
        return target in text
    if mode == "3":
        return text != target
    if mode == "4":
        return text == ""
    if mode == "5":
        return text != ""
    return False


def _rule_desc(rule):
    """生成规则的可读描述。"""
    mode = rule["mode"]
    label = _MODE_LABEL.get(mode, mode)
    target = rule.get("target")
    tail = f"「{target}」" if target is not None else ""
    return f"{label}{tail}"


# ---------------------------------------------------------------- 源文件收集

def _list_dir_files(directory, exts=(".xlsx", ".xlsm")):
    """列出目录下指定后缀的文件。"""
    try:
        return sorted(
            f for f in os.listdir(directory)
            if f.lower().endswith(exts) and not f.startswith("~$")
        )
    except OSError:
        return []


def _browse_for_file():
    """浏览常见目录挑选一个文件（仅 Android 模式可用）。返回绝对路径或 None。"""
    dirs = paths.common_dirs()
    pick = ui.menu(
        [(str(i + 1), d) for i, d in enumerate(dirs)],
        title_text="选择目录",
        allow_back=True,
    )
    if pick is None:
        return None
    directory = dirs[int(pick) - 1]
    if not os.path.isdir(directory):
        ui.error(f"目录不存在：{directory}")
        return None
    files = _list_dir_files(directory)
    if not files:
        ui.warn(f"{directory} 中没有 xlsx 文件。")
        return None
    fpick = ui.menu(
        [(str(i + 1), f) for i, f in enumerate(files)],
        title_text=f"{directory} 中的表格",
        allow_back=True,
    )
    if fpick is None:
        return None
    return os.path.join(directory, files[int(fpick) - 1])


def _add_file_manual(files):
    """手动输入一个或多个文件路径（分号/换行分隔也可以逐个输入）。"""
    raw = ui.ask("文件路径（直接回车取消）")
    p = paths.normalize_path(raw)
    if not p:
        return False
    if not os.path.isfile(p):
        ui.error("文件不存在。")
        return False
    if not paths.is_readable_file(p):
        ui.error("文件不可读，请检查权限。")
        return False
    if not p.lower().endswith((".xlsx", ".xlsm")):
        ui.warn("该文件不是 .xlsx/.xlsm，可能无法解析。")
        if not ui.ask_yes_no("仍要添加吗", "n"):
            return False
    ap = os.path.abspath(p)
    if ap in files:
        ui.warn("该文件已在列表中。")
        return False
    files.append(ap)
    ui.success(f"已添加：{ap}")
    return True


def _add_files_by_paste(files):
    """批量粘贴多个路径，支持换行 / 分号分隔。"""
    print("  每行一个路径，或用分号分隔；输入空行结束。")
    added = 0
    while True:
        raw = ui.ask("路径（空行结束）")
        if raw == "":
            break
        for chunk in raw.replace(";", "\n").split("\n"):
            p = paths.normalize_path(chunk)
            if not p:
                continue
            if not os.path.isfile(p):
                ui.error(f"不存在，跳过：{p}")
                continue
            ap = os.path.abspath(p)
            if ap in files:
                ui.warn(f"已在列表中，跳过：{ap}")
                continue
            files.append(ap)
            ui.success(f"已添加：{ap}")
            added += 1
    return added > 0


def _add_files_by_dir(files):
    """把某个目录下的全部 xlsx 加入列表（支持序号范围选择）。

    Android 模式下先列出常见目录供选择，并保留「手动输入目录路径」；
    Windows 模式下直接要求输入目录路径。
    """
    dirs = paths.common_dirs()
    if dirs:
        options = [(str(i + 1), d) for i, d in enumerate(dirs)]
        options.append((str(len(options) + 1), "手动输入目录路径"))
        pick = ui.menu(options, title_text="选择目录", allow_back=True)
        if pick is None:
            return False
        idx = int(pick) - 1
        if idx == len(dirs):
            directory = paths.normalize_path(ui.ask("目录路径（直接回车取消）"))
        else:
            directory = dirs[idx]
    else:
        directory = paths.normalize_path(ui.ask("目录路径（直接回车取消）"))

    if not directory:
        return False
    if not os.path.isdir(directory):
        ui.error(f"目录不存在：{directory}")
        return False

    names = _list_dir_files(directory)
    if not names:
        ui.warn(f"{directory} 中没有 xlsx 文件。")
        return False

    print()
    print("  目录中的表格：")
    for i, n in enumerate(names):
        print(f"    [{i + 1}] {n}")
    print("  可输入如 1-3,5 的序号范围；直接回车表示全部。")

    raw = ui.ask("选择序号", default="all")
    if raw.lower() in ("all", ""):
        idxs = list(range(1, len(names) + 1))
    else:
        idxs = _parse_index_spec(raw, len(names))
        if not idxs:
            ui.error("序号无效。")
            return False

    added = 0
    for i in idxs:
        ap = os.path.abspath(os.path.join(directory, names[i - 1]))
        if ap in files:
            continue
        files.append(ap)
        added += 1
    ui.success(f"从 {directory} 添加了 {added} 个文件。")
    return added > 0


def _collect_files():
    """收集源文件列表（可多个）。返回绝对路径列表或 None（取消）。"""
    ui.subtitle("步骤 1 / 6：导入 xlsx 文件（可多个）")
    files = []
    while True:
        print()
        if files:
            print("  当前文件列表：")
            for i, f in enumerate(files):
                print(f"    [{i + 1}] {f}")
        else:
            print("  （尚未添加任何文件）")

        choice = ui.menu(
            [
                ("1", "添加单个文件（浏览目录 / 手动输入）"
                      if paths.has_common_dirs() else "添加单个文件（手动输入路径）"),
                ("2", "批量粘贴多个路径"),
                ("3", "添加整个目录的表格"),
                ("4", "移除某个文件"),
                ("5", "清空列表"),
                ("6", "完成，进入下一步"),
            ],
            title_text="文件列表操作",
            allow_back=True,
        )
        if choice is None:
            return None

        if choice == "1":
            # 「浏览常见目录」只在 Android 模式提供：Windows 下不存在 /sdcard 等目录
            if paths.has_common_dirs() and ui.ask_yes_no("是否浏览常见目录", "n"):
                p = _browse_for_file()
                if p:
                    ap = os.path.abspath(p)
                    if ap in files:
                        ui.warn("该文件已在列表中。")
                    else:
                        files.append(ap)
                        ui.success(f"已添加：{ap}")
            else:
                _add_file_manual(files)
        elif choice == "2":
            _add_files_by_paste(files)
        elif choice == "3":
            _add_files_by_dir(files)
        elif choice == "4":
            if not files:
                ui.warn("列表为空。")
                continue
            raw = ui.ask("要移除的序号（如 1 或 1,3）")
            idxs = _parse_index_spec(raw, len(files))
            if not idxs:
                ui.error("序号无效。")
                continue
            for i in sorted(idxs, reverse=True):
                removed = files.pop(i - 1)
                ui.info(f"已移除：{removed}")
        elif choice == "5":
            if files and ui.ask_yes_no("确认清空列表", "n"):
                files.clear()
                ui.info("已清空。")
        elif choice == "6":
            if not files:
                ui.warn("请至少添加一个文件。")
                continue
            return files


# ---------------------------------------------------------------- 工作表选择

def _choose_sheets(wb, file_path):
    """为某个工作簿选择工作表（多选 / 全部）。返回名称列表或 None。"""
    names = wb.sheetnames
    print()
    print(f"  {os.path.basename(file_path)} 的工作表：")
    for i, name in enumerate(names):
        ws = wb[name]
        print(f"    [{i + 1}] {name}  "
              f"（{ws.max_row} 行 × {ws.max_column} 列）")

    raw = ui.ask("选择工作表序号（回车=全部，如 1-2,4）", default="all")
    if raw.lower() in ("all", ""):
        return list(names)
    idxs = _parse_index_spec(raw, len(names))
    if not idxs:
        ui.error("序号无效，已跳过该文件。")
        return None
    return [names[i - 1] for i in idxs]


# ---------------------------------------------------------------- 表头定位

def _ask_header_rows(max_row):
    """询问表头占用哪些行，支持多行表头。返回行号列表。"""
    start = ui.ask_int("表头起始行", default=1, min_value=1,
                       max_value=max(1, max_row))
    if start is None:
        start = 1
    end = ui.ask_int("表头结束行", default=start, min_value=start,
                     max_value=max(1, max_row))
    if end is None:
        end = start
    return list(range(start, end + 1))


# ---------------------------------------------------------------- 条件设置

def _ask_match_mode(title="选择匹配方式"):
    pick = ui.menu(MATCH_MODES, title_text=title, allow_back=True)
    return pick


def _build_rule(col_index, col_name, col_letter, mode, target):
    return {
        "col_index": col_index,
        "col_name": col_name,
        "col_letter": col_letter,
        "mode": mode,
        "target": target,
    }


def _set_row_rules(cols):
    """设置行条件规则列表（可为空）。cols: [(列序号, 列名), ...]"""
    rules = []
    while True:
        print()
        if rules:
            print("  当前行条件（多个条件之间为「同时满足」）：")
            for i, r in enumerate(rules):
                print(f"    [{i + 1}] {r['col_name']}（{r['col_letter']} 列） "
                      f"{_rule_desc(r)}")
        else:
            print("  （尚未设置行条件，若无行条件则保留全部数据行）")

        choice = ui.menu(
            [
                ("1", "添加一条行条件"),
                ("2", "移除某条条件"),
                ("3", "清空行条件"),
                ("4", "完成行条件设置"),
            ],
            title_text="行条件设置",
            allow_back=True,
        )
        if choice is None:
            return None

        if choice == "1":
            col_items = [
                (str(i + 1), name, f"第 {_col_letter(c)} 列")
                for i, (c, name) in enumerate(cols)
            ]
            pick = ui.menu(col_items, title_text="按哪一列判断", allow_back=True)
            if pick is None:
                continue
            col_index, col_name = cols[int(pick) - 1]
            mode = _ask_match_mode()
            if mode is None:
                continue
            target = None
            if mode not in ("4", "5"):
                target = _cell_text(ui.ask("目标值（例如 15）"))
                if target == "":
                    ui.warn("目标值为空，已取消该条件。")
                    continue
            rules.append(_build_rule(
                col_index, col_name, _col_letter(col_index), mode, target))
            ui.success(f"已添加行条件：{col_name} {_rule_desc(rules[-1])}")
        elif choice == "2":
            if not rules:
                ui.warn("没有可移除的条件。")
                continue
            raw = ui.ask("要移除的序号（如 1 或 1,3）")
            idxs = _parse_index_spec(raw, len(rules))
            if not idxs:
                ui.error("序号无效。")
                continue
            for i in sorted(idxs, reverse=True):
                rules.pop(i - 1)
            ui.info("已移除。")
        elif choice == "3":
            if rules and ui.ask_yes_no("确认清空行条件", "n"):
                rules.clear()
                ui.info("已清空。")
        elif choice == "4":
            return rules


def _set_col_rules(cols):
    """设置列条件（列名筛选），规则作用于列名。

    返回满足条件的列序号列表（保持原顺序）；无规则表示保留全部列。
    """
    rules = []
    while True:
        print()
        if rules:
            print("  当前列条件（针对列名，多个条件之间为「同时满足」）：")
            for i, r in enumerate(rules):
                print(f"    [{i + 1}] 列名 {_rule_desc(r)}")
        else:
            print("  （尚未设置列条件，将保留所有列）")

        choice = ui.menu(
            [
                ("1", "添加一条列条件（按列名匹配）"),
                ("2", "移除某条条件"),
                ("3", "清空列条件"),
                ("4", "完成列条件设置"),
            ],
            title_text="列条件设置",
            allow_back=True,
        )
        if choice is None:
            return None

        if choice == "1":
            mode = _ask_match_mode("选择列名匹配方式")
            if mode is None:
                continue
            target = None
            if mode not in ("4", "5"):
                target = _cell_text(ui.ask("列名关键词（例如 15）"))
                if target == "":
                    ui.warn("目标值为空，已取消该条件。")
                    continue
            rules.append({"mode": mode, "target": target})
            ui.success(f"已添加列条件：列名 {_rule_desc(rules[-1])}")
        elif choice == "2":
            if not rules:
                ui.warn("没有可移除的条件。")
                continue
            raw = ui.ask("要移除的序号（如 1 或 1,3）")
            idxs = _parse_index_spec(raw, len(rules))
            if not idxs:
                ui.error("序号无效。")
                continue
            for i in sorted(idxs, reverse=True):
                rules.pop(i - 1)
            ui.info("已移除。")
        elif choice == "3":
            if rules and ui.ask_yes_no("确认清空列条件", "n"):
                rules.clear()
                ui.info("已清空。")
        elif choice == "4":
            if not rules:
                return list(range(1, len(cols) + 1))  # 全部列
            keep = []
            for c, name in cols:
                if all(_match(name, r["mode"], r["target"]) for r in rules):
                    keep.append(c)
            return keep


# ---------------------------------------------------------------- 提取核心

def _extract_sheet(ws, header_rows, row_rules, keep_cols):
    """从单个工作表提取数据。

    返回 (命中数据行列表, 命中行号列表)。
    row_rules 为空表示保留全部数据行；keep_cols 为空表示无列。
    """
    data_start = header_rows[-1] + 1
    data_rows = []
    hit_rows = []
    for r in range(data_start, ws.max_row + 1):
        ok = True
        for rule in row_rules:
            text = _cell_text(ws.cell(row=r, column=rule["col_index"]).value)
            if not _match(text, rule["mode"], rule["target"]):
                ok = False
                break
        if ok:
            data_rows.append([ws.cell(row=r, column=c).value for c in keep_cols])
            hit_rows.append(r)
    return data_rows, hit_rows


# ---------------------------------------------------------------- 主流程

def run():
    """工具入口。"""
    openpyxl = _load_openpyxl()
    if openpyxl is None:
        ui.pause()
        return

    # 1. 收集文件
    files = _collect_files()
    if not files:
        return

    # 2. 打开各文件并逐文件选择工作表
    ui.subtitle("步骤 2 / 6：选择工作表")
    plan = []  # [(file_path, wb, [sheet_name, ...]), ...]
    for fp in files:
        try:
            wb = openpyxl.load_workbook(fp, data_only=True, read_only=False)
        except Exception as exc:  # noqa: BLE001
            ui.error(f"打开失败，已跳过 {fp}：{exc}")
            continue
        sheets = _choose_sheets(wb, fp)
        if not sheets:
            ui.warn(f"{os.path.basename(fp)}：未选择工作表，已跳过。")
            continue
        plan.append((fp, wb, sheets))

    if not plan:
        ui.error("没有任何可处理的工作表。")
        ui.pause()
        return

    # 3. 表头行（以第一个工作表的规模提示，行数为通用询问）
    ui.subtitle("步骤 3 / 6：定位表头")
    print("  单行表头：直接回车（默认第 1 行）。")
    print("  多行表头：输入起止行号，例如起 1 止 2，表示表头占第 1~2 行。")
    first_ws = plan[0][1][plan[0][2][0]]
    header_rows = _ask_header_rows(first_ws.max_row)
    header_note = (
        f"第 {header_rows[0]} 行"
        if len(header_rows) == 1
        else f"第 {header_rows[0]}~{header_rows[-1]} 行（共 {len(header_rows)} 行）"
    )
    ui.info(f"表头范围：{header_note}")

    # 4. 行条件（基于第一个工作表读到的列）
    ui.subtitle("步骤 4 / 6：设置行条件")
    first_cols = _read_header(first_ws, header_rows)
    if not first_cols:
        ui.error("未在所选表头行读取到任何列名，请确认表头行范围。")
        ui.pause()
        return
    print("  以下列名取自第一个工作表，用于设置条件：")
    for i, (c, name) in enumerate(first_cols):
        print(f"    [{i + 1}] {name}  （{_col_letter(c)} 列）")
    row_rules = _set_row_rules(first_cols)
    if row_rules is None:
        return

    # 5. 列条件
    ui.subtitle("步骤 5 / 6：设置列条件")
    keep_cols = _set_col_rules(first_cols)
    if keep_cols is None:
        return
    if not keep_cols:
        ui.warn("列条件没有匹配到任何列，无法导出。")
        ui.pause()
        return

    # 开始按工作簿/工作表逐个提取
    results = []  # [(file_path, sheet_name, header_rows_data, data_rows, hit_count)]
    total_hits = 0
    for fp, wb, sheets in plan:
        for sname in sheets:
            ws = wb[sname]
            cols = _read_header(ws, header_rows)
            if not cols:
                ui.warn(f"{os.path.basename(fp)} / {sname}：未读到表头，已跳过。")
                continue
            # 按列序号对齐的列名映射
            name_by_col = {c: name for c, name in cols}
            # 允许 keep_cols 中的列在该表不存在时跳过
            use_cols = [c for c in keep_cols if c in name_by_col]
            if not use_cols:
                ui.warn(f"{os.path.basename(fp)} / {sname}：无匹配的列，已跳过。")
                continue
            data_rows, hit_rows = _extract_sheet(ws, header_rows, row_rules, use_cols)
            if not data_rows:
                continue
            # 表头数据（按 use_cols 取列名）
            merged = _merged_value_map(ws)
            header_data = []
            for hr in header_rows:
                header_data.append([
                    _header_cell_text(ws, merged, hr, c) for c in use_cols
                ])
            results.append((fp, sname, header_data, data_rows, len(hit_rows)))
            total_hits += len(hit_rows)

    # 预览
    ui.subtitle("步骤 6 / 6：预览与导出")
    if not results:
        ui.warn("没有任何数据符合条件，未生成文件。")
        ui.pause()
        return
    print()
    print(f"  命中汇总（共 {len(results)} 个工作表 / {total_hits} 行）：")
    for i, (fp, sname, _, drows, hits) in enumerate(results):
        print(f"    [{i + 1}] {os.path.basename(fp)} / {sname}：{hits} 行")

    print()
    print("  行条件：" + ("无（保留全部数据行）" if not row_rules else ""))
    for r in row_rules:
        print(f"    - {r['col_name']}（{r['col_letter']} 列） {_rule_desc(r)}")
    kept_names = " / ".join(_col_letter(c) for c in keep_cols)
    ui.info(f"保留列：{kept_names}（共 {len(keep_cols)} 列）")

    # 导出
    default_name = paths.default_output_name(files[0], "合并提取")
    out_path = ui.ask("输出文件路径",
                      default=os.path.join(paths.ensure_output_dir(), default_name))
    out_path = paths.normalize_path(out_path)
    if not out_path:
        ui.warn("输出路径为空，已取消。")
        return
    if not out_path.lower().endswith((".xlsx", ".xlsm")):
        out_path += ".xlsx"
    out_dir = os.path.dirname(os.path.abspath(out_path))
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError as exc:
        ui.error(f"无法创建输出目录：{exc}")
        ui.pause()
        return

    try:
        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)  # 移除默认空表

        used_titles = set()
        for fp, sname, header_data, data_rows, _ in results:
            title = _unique_sheet_title(sname, used_titles)
            out_ws = out_wb.create_sheet(title=title)
            hr_count = len(header_data)
            for i, hrow in enumerate(header_data):
                for j, val in enumerate(hrow):
                    out_ws.cell(row=i + 1, column=j + 1).value = val
            for i, drow in enumerate(data_rows):
                for j, val in enumerate(drow):
                    out_ws.cell(row=hr_count + 1 + i, column=j + 1).value = val
        out_wb.save(out_path)
    except PermissionError:
        ui.error("写入被拒绝，目标文件可能正被占用，或路径无写权限。")
        ui.pause()
        return
    except Exception as exc:  # noqa: BLE001
        ui.error(f"导出失败：{exc}")
        ui.pause()
        return

    print()
    ui.success(f"已生成：{out_path}")
    ui.info(f"共写入 {len(results)} 个工作表 / {total_hits} 行数据。")

    if ui.ask_yes_no("是否继续处理其他文件", "n"):
        run()
        return
    ui.pause()


def _unique_sheet_title(name, used):
    """生成不重复的工作表名（Excel 限制 31 字符）。"""
    base = (name or "结果")[:31] or "结果"
    title = base
    n = 2
    while title in used:
        suffix = f"_{n}"
        title = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(title)
    return title
