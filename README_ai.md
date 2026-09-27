# ToolsBox 技术手册（面向 AI Agent / 开发者）

> **本文件的用途**：供 AI coding agent 与开发者阅读的技术说明。它包含完整实现细节、
> 平台差异、模块职责、依赖方向与编码约定，可以作为改动代码前的上下文。
>
> 面向普通用户、准备放到 GitHub 首页的介绍请看 [`README.md`](README.md)。
> 两份文件描述同一套代码：**改动代码后请同步更新本文档。**

一个用 Python 编写的**命令行**工具箱。同一套代码同时支持 Windows 与 Android(Termux)：
启动时先询问运行平台，然后平铺列出全部工具，输入序号即可执行。

---

## 0. 给 AI Agent 的快速摘要

改代码前只需知道这几件事：

- **入口**：`main.py` → 先定平台（`core/runtime.py`）→ 再进工具菜单（`core/registry.py`）。
- **平台差异只有三处**：清屏命令、是否提供「浏览常见目录」、提示里的路径示例。
  全部集中在 `core/runtime.py` 与 `core/paths.py`，**不要在工具里写 `os.name` 判断**。
- **依赖方向是单向的**：`tools → paths / ui → runtime`。
  `core/runtime.py` 顶层不导入任何项目内模块；`runtime.ask_platform()` 用函数内延迟导入拿 `ui`。
  改动时不要破坏这个方向，否则会循环导入。
- **输出必须是纯文本**：不要引入 ANSI 颜色转义序列（旧版 Windows cmd 会显示成乱码）。
- **读输入的两条铁律**：
  1. 在 `while True` 循环里读输入必须用 `ui.ask_raw()` 并处理 `None`（表示 EOF），
     否则会无限刷「无效的序号」；普通一次性输入用 `ui.ask()` 即可。
  2. 不要用 `ui.ask()` 的返回值去区分「用户回车」和「输入流结束」——它区分不了。
- **工具入口约定**：每个 `tools/<name>/__init__.py` 暴露 `NAME` / `ORDER` / `DESCRIPTION` / `run`，
  并在 `core/registry.py` 的 `TOOLS` 里登记。详见第 11 节。
- **许可证是 Unlicense（公有领域）**：项目**放弃全部著作权**。
  不要给源码加版权头、`SPDX-License-Identifier` 或其他许可证标识，也不要「顺手」改成 MIT 等协议。
- **不要动** `.backup/legacy-android/`：那是移植前的存档，不参与运行。

---

## 1. 环境要求

| 项目 | 要求 |
| --- | --- |
| 系统 | Windows 10 / 11；Android + Termux |
| Python | 3.8 或更高（Windows 验证环境：3.13.5） |
| 依赖库 | `openpyxl`（读写 xlsx）、`Pillow`（图片转 PDF） |

`python --version` 可查看当前版本。若 Windows 下提示找不到 `python`，
请重装 Python 并勾选 **Add Python to PATH**，或改用 `py` 命令。

---

## 2. 安装

**Windows**

```bat
cd /d E:\ouroooo\AIwork-space\ToolsBox
python -m pip install -r requirements.txt
```

**Android / Termux**

```bash
pkg install python
cd /storage/emulated/0/QUJIProject/AIwork-space/ToolsBox
python -m pip install -r requirements.txt
```

`requirements.txt` 内容：

```
openpyxl>=3.0
Pillow>=9.0
```

> Termux 下若要读写手机共享存储（如 `/sdcard/Download`），
> 需先执行一次 `termux-setup-storage` 并授予存储权限。

---

## 3. 运行

两个平台命令相同：

```bash
python main.py
```

启动后会先询问运行平台：

```
运行平台
--------------------------------------------------------
  本工具箱同时支持 Windows 与 Android，请确认当前运行平台。
  两种模式的差别：清屏命令、是否提供「浏览常见目录」、提示里的路径示例。

--------------------------------------------------------
    [1] Windows
        使用 cls 清屏；路径形如 C:\Users\Public\Documents\成绩表.xlsx
    [2] Android / Termux
        使用 clear 清屏；可浏览 /sdcard、/storage/emulated/0 等常见目录
--------------------------------------------------------
请输入序号（默认 1）:
```

- **直接回车**即采用自动探测结果：Windows 用 `os.name` 判断，
  其余环境（Android/Termux、Linux、macOS）统一按 Android 处理 —— 它们清屏命令同为 `clear`。
- 平台选定后，标题栏会显示出来，例如 `版本 0.2.0 · 平台 Windows · 输入序号选择工具 · 输入 0 退出`。

### 跳过询问（可选）

自动化脚本或桌面快捷方式里可以直接指定平台，跳过交互：

```bash
python main.py --platform=windows
python main.py --platform=android
```

取值无法识别时（例如 `--platform=foo`）会提示并回退到交互询问，不会静默忽略。

> ⚠️ **Windows 下请勿双击 `main.py`。** 双击时标准输入会立刻结束（EOF），
> 程序会采用默认平台后立即退出。请在 cmd 或 PowerShell 窗口中运行。

### 交互约定

- 菜单里输入 `0` 返回上一级；主菜单输入 `0` 退出程序。
  例外：「图片转 PDF」的**尺寸模式**菜单没有返回项，必须从三项中选一项。
- 提示符处按 `Ctrl+C` 等同于「取消当前这一步」。
- 输入流意外结束（EOF）时不会死循环，会按「取消」逐层干净退出。

---

## 4. 两种模式的差别

平台差异全部集中在 `core/runtime.py` 与 `core/paths.py`：

| 项目 | Windows 模式 | Android / Termux 模式 |
| --- | --- | --- |
| 清屏命令 | `cls` | `clear` |
| 「浏览常见目录」 | **不提供**（Windows 下不存在这些目录） | 提供以下 6 个：`/sdcard/Download`、`/sdcard/Documents`、`/storage/emulated/0/Download`、`/storage/emulated/0/Documents`、`/sdcard`、`/storage/emulated/0` |
| 提示里的路径示例 | `C:\Users\Public\Documents\成绩表.xlsx` | `/sdcard/Download/成绩表.xlsx` |

除以上三点外，两个平台行为完全一致，且都是**纯文本输出，不使用 ANSI 颜色转义序列**。

---

## 5. 路径输入

**Windows 模式**

- 路径含空格时用英文双引号包住，程序会自动去掉引号：
  `"C:\Users\Public\My Documents\成绩表.xlsx"`
- 反斜杠 `\` 与正斜杠 `/` 都能识别，下面两种写法等价：
  `C:\Users\Public\成绩表.xlsx` / `C:/Users/Public/成绩表.xlsx`
- 支持 `~` 代表当前用户目录，例如 `~\Desktop\成绩表.xlsx`。
- 在资源管理器里 `Shift + 右键` →「复制文件地址」可得到完整路径，直接粘贴即可。

**Android 模式**

- 共享存储路径形如 `/sdcard/Download/成绩表.xlsx`，可直接输入，
  也可用「浏览常见目录」从上面 6 个目录里挑选文件。
- 输出目录 `output/` 不存在时会自动创建。

---

## 6. 控制台中文显示（Windows）

- 在 cmd / PowerShell 窗口**交互运行**时中文显示正常
  （Python 3.6+ 使用宽字符控制台 API，不受代码页影响）。
- 但若把输出**重定向到文件或用管道**传给其它程序，Python 会按系统区域编码写出
  （简体中文 Windows 为 GBK / cp936），用 UTF-8 编辑器打开会乱码。需要统一为 UTF-8 时：

  ```bat
  chcp 65001
  set PYTHONIOENCODING=utf-8
  python main.py
  ```

---

## 7. 目录结构

```
ToolsBox/
├── main.py                      # 程序入口：询问平台 → 工具选择菜单
├── requirements.txt
├── README.md                    # 面向用户 / GitHub 首页的介绍
├── README_ai.md                 # 本文件：面向 AI Agent / 开发者的技术手册
├── UNLICENSE                    # The Unlicense（公有领域，放弃著作权）
├── core/                        # 公共基础设施
│   ├── __init__.py
│   ├── runtime.py               # 运行平台：探测 / 询问 / 切换、清屏命令
│   ├── ui.py                    # 终端交互：标题、菜单、输入、提示
│   ├── paths.py                 # 路径：输出目录、常见目录、路径示例
│   └── registry.py              # 工具注册表
├── tools/                       # 各工具独立文件夹
│   ├── excel_extract/           # 【工具 1】xlsx 数据提取
│   ├── image_to_pdf/            # 【工具 2】图片转 PDF
│   └── excel_multi_extract/     # 【工具 3】多表行列提取
├── output/                      # 默认输出目录（自动创建）
└── .backup/legacy-android/      # 移植前的 Android 专用版本存档（不参与运行）
```

---

## 8. 工具说明

### 工具 1：xlsx 数据提取

用途：从 xlsx 中提取「某列等于 / 包含某内容」的所有行（例如班级 = 15），生成新的 xlsx。

| 步骤 | 内容 |
| --- | --- |
| 1 | 选择 xlsx 文件：输入完整路径；**Android 模式**下还可先浏览常见目录挑选 |
| 2 | 选择工作表 |
| 3 | 指定表头行范围：单行表头直接回车（默认第 1 行）；多行表头输入起止行，如 1~2 |
| 4 | 选择筛选列（读取多行表头并拼接列名，兼容合并单元格） |
| 5 | 选择匹配方式：精确 / 包含 / 不等于 / 为空 / 非空，并输入目标值 |
| 6 | 预览命中行数与行号，确认后导出到新 xlsx |

- 输出默认保存到 `output/`，文件名形如 `成绩表_班级_15.xlsx`。
- 导出后可选择在结果文件上再次筛选。

### 工具 2：图片转 PDF

用途：把一张或多张图片合并成一个 PDF。

| 步骤 | 内容 |
| --- | --- |
| 1 | 收集图片：添加单张（**Android 模式**可浏览目录）、批量粘贴路径、添加整个目录，并可调整顺序 / 删除 / 清空 |
| 2 | 选择尺寸模式：原图尺寸 / A4 自适应 / 固定宽度（自定义 mm） |
| 3 | 选择输出位置（默认 `output/第一张图片名_合并.pdf`） |
| 4 | 确认后导出 |

- 支持格式：`jpg` `jpeg` `png`，另兼容 `bmp` `gif` `webp` `tif` `tiff`（取决于 Pillow 解码能力）。
- 带透明通道的图片会以白色为底填充，避免转 PDF 时出现黑块。
- 输出文件已存在时会询问是否覆盖，选否则自动加序号（`_2`、`_3`…）。

### 工具 3：多表行列提取

用途：一次导入**多个 xlsx** 或**一个 xlsx 的多个工作表**，对每个工作表按「行条件」和/或「列条件」
（至少设置一个）提取数据，并把所有结果合并到**一个新的 xlsx**（每个工作表对应输出文件里的一个 sheet）。

典型场景：

- 提取所有「班级列包含 15」的数据行（行条件）。
- 提取所有列名「包含 15」的整列（列条件）。
- 两者同时使用则取交集：命中的行中只保留命中的列。

| 步骤 | 内容 |
| --- | --- |
| 1 | 添加待处理文件：添加单个文件（**Android 模式**可浏览常见目录，两平台都可手动输入）、批量粘贴多个路径（换行或分号分隔）、添加整个目录（**Android 模式**可先从常见目录里挑），可移除 / 清空 |
| 2 | 逐文件选择工作表：直接回车 = 全部工作表；也可输入 `1-2,4` 选择部分 |
| 3 | 定位表头行范围：单行表头直接回车（默认第 1 行）；多行表头输入起止行，如 1~2 |
| 4 | 设置**行条件**（可多条，之间为 AND 关系）：选择列 + 匹配方式 + 目标值 |
| 5 | 设置**列条件**（可多条）：按列名匹配需保留的列；不设置则保留全部列 |
| 6 | 预览各工作表命中行数，确认后合并导出 |

说明：

- 行条件与列条件**至少需设置一个**；两者都设置时按交集处理。
- 未命中任何列（列条件全不匹配）会提示并退出，避免生成空表。
- 输出默认保存到 `output/`，文件名形如 `原文件名_合并提取.xlsx`；
  不同工作表输出为不同 sheet，重名 sheet 自动加序号。
- 序号输入支持 `1-3,5` 这种写法，中英文逗号、空格都能识别。

---

## 9. 双平台是怎么实现的

### 9.1 平台状态：`core/runtime.py`

这个模块持有「当前平台」这一个状态，并提供平台相关的行为：

| 函数 | 作用 |
| --- | --- |
| `detect()` | 自动探测：`os.name == "nt"` → Windows，否则 Android |
| `ask_platform()` | 启动时询问用户，默认值取自 `detect()`，回车即采用 |
| `current()` / `set_platform()` | 读写当前平台 |
| `is_windows()` / `is_android()` | 供其它模块判断 |
| `label()` | 中文显示名，如 `Android / Termux` |
| `clear_command()` | 返回 `cls` 或 `clear` |
| `parse_platform_arg()` | 归一化 `--platform=` 的取值 |

启动流程在 `main.py`：`decide_platform()` 先看命令行有没有 `--platform=`，
有且合法就直接采用、不询问；否则调用 `runtime.ask_platform()` 交互询问。

### 9.2 平台相关的路径：`core/paths.py`

- `common_dirs()` —— Android 模式返回 6 个 `/sdcard` 等目录；Windows 模式返回空列表。
- `has_common_dirs()` —— 各工具据此决定是否显示「浏览常见目录」入口。
- `example_path()` —— 交互提示里的示例路径，随平台变化。

各工具的写法都是同一个模式：

```python
if paths.has_common_dirs() and ui.ask_yes_no("是否浏览常见目录", "n"):
    ...  # 仅 Android 模式会进入这里
```

因此 Windows 模式下这套代码路径**完全不会被执行**，也不会出现空菜单或「目录不存在」。

### 9.3 为什么不会循环导入

`core/runtime.py` 在**顶层不导入任何项目内模块**，所以：

- `core/ui.py` 可以放心地在顶层 `from core import runtime`（`clear_screen()` 要用 `clear_command()`）；
- `core/paths.py` 也可以顶层 `from core import runtime`；
- 只有 `runtime.ask_platform()` 需要用到 `ui`，它采用**函数内延迟导入**来避开循环。

依赖方向是单向的：`tools → paths / ui → runtime`。

### 9.4 关于 `[android]` 注释标记

早期「只保 Windows」阶段，代码里用 `[android]` 注释标记被注释掉的 Android 代码。
现在已改为一套代码同时支持两平台，这些标记已全部清理，不再需要。

---

## 10. 存档文件

`.backup/legacy-android/` 存放移植前的 Android 专用版本，**不参与运行**，仅供对照：

| 文件 | 说明 |
| --- | --- |
| `main-android.py` | 原入口文件（移植前），引用了已删除的 `ui.C` 颜色常量，直接运行会报错 |
| `excel_extract.extractor.py.bak` | 原「工具 1」实现，内含早先的 `COMMON_DIRS` 与浏览逻辑 |
| `README_android.md` | 原 Android / Termux 专用说明（内容已并入本文件） |

---

## 11. 新增工具的方法

1. 在 `tools/` 下新建文件夹，例如 `tools/my_tool/`；
2. 在文件夹内实现逻辑，并在 `__init__.py` 中暴露入口：

   ```python
   NAME = "工具名"
   ORDER = 1                      # 数字越小越靠前，默认 99
   DESCRIPTION = "简述"
   from .impl import run          # run() 为入口，自行完成交互
   ```

3. 在 `core/registry.py` 的 `TOOLS` 列表中加入该模块即可。

编写交互时请复用 `core/ui.py` 与 `core/paths.py`：

```python
from core import ui, paths

def run():
    ui.title("我的工具")
    name = ui.ask("请输入名称")                 # 文本输入
    n    = ui.ask_int("数量", default=1)        # 整数输入（带范围校验）
    if ui.ask_yes_no("确认", "y"):              # 是/否
        ...
    path = ui.menu([("1", "选项一"), ("2", "选项二")], title_text="请选择")
    ui.success("完成")
    ui.pause()                                  # 按回车返回主菜单
```

需要处理平台差异时：

```python
from core import runtime, paths

if runtime.is_windows():
    ...
if paths.has_common_dirs():        # 仅 Android 模式为真
    ...                            # 提供「浏览常见目录」之类的快捷入口
```

> 两个注意点：
> 1. 若在 `while True` 循环里读输入，请用 `ui.ask_raw()` 并处理 `None`
>    （表示输入流已结束），否则会死循环；普通场景直接用 `ui.ask()` 即可。
> 2. 输出请不要使用 ANSI 颜色转义序列，保持纯文本。
