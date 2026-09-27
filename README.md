<div align="center">

# ToolsBox 多功能工具箱

**一个用 Python 写的命令行工具箱 —— 一套代码同时支持 Windows 与 Android (Termux)**

[![License](https://img.shields.io/badge/License-Unlicense-blue.svg)](UNLICENSE)
[![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Android-lightgrey.svg)](#-快速开始)
[![Dependencies](https://img.shields.io/badge/dependencies-openpyxl%20%7C%20Pillow%20%7C%20pypdf-green.svg)](requirements.txt)

</div>

---

## 简介

ToolsBox 把几个日常处理表格和图片的小需求，做成了一个开箱即用的命令行工具箱：

- 📊 **xlsx 数据提取** —— 按某列条件（如「班级 = 15」）筛选后导出新表
- 🖼️ **图片转 PDF** —— 多张 jpg / png 合并成一个 PDF，可选 A4 自适应
- 📚 **多表行列提取** —— 多文件 / 多工作表按行、列条件提取并合并到一个新表
- 📄 **PDF 单页提取** —— 从 PDF 中取出指定的一页，另存为新的 PDF

不需要 GUI、不需要数据库、不联网。启动后输入序号即可使用。

## ✨ 特性

| | |
| --- | --- |
| 🪟🤖 **双平台单代码库** | 同一套代码跑在 Windows 与 Android (Termux) 上，启动时确认平台即可，差异（清屏命令、常见目录、路径示例）全部自动切换 |
| 🧭 **零学习成本** | 全程菜单式交互，每步都有提示和预览；序号支持 `1-3,5` 这种范围写法 |
| 🪶 **依赖极少** | 只需要 `openpyxl`、`Pillow` 与 `pypdf`，其余全部是标准库 |
| 🧩 **易于扩展** | 新增一个工具只需建一个文件夹 + 在注册表登记一行，互不影响 |
| 🖥️ **终端友好** | 输出为纯文本，不使用 ANSI 转义序列，老旧 Windows cmd 下也不会满屏乱码 |
| 🛡️ **不易崩** | 单个工具出错会被捕获并返回主菜单；输入流意外结束（EOF）也会干净退出而不是死循环 |

## 📸 界面预览

```
========================================================
                       ToolsBox
========================================================
  版本 0.3.0  ·  平台 Windows  ·  输入序号选择工具  ·  输入 0 退出

请选择工具
--------------------------------------------------------
  [1] xlsx 数据提取  按列筛选（如班级=15）后导出新表
  [2] 图片转 PDF  单个或多个 jpg/png/jpeg 合并为 PDF
  [3] 多表行列提取  多 xlsx / 多工作表按行或列条件提取并合并
  [4] PDF 单页提取  把 PDF 中的指定一页另存为新的 PDF
  [0] 退出
--------------------------------------------------------
请输入序号: 1

正在运行：xlsx 数据提取
--------------------------------------------------------
步骤 5 / 6：预览结果
[信息] 文件：C:\Users\Public\Documents\成绩表.xlsx
[信息] 工作表：成绩
[信息] 筛选条件：班级（A 列） 包含关键词「15」
[信息] 数据范围：第 2 ~ 421 行，共 37 行命中。
[成功] 已生成：E:\ToolsBox\output\成绩表_班级_15.xlsx
```

## 🚀 快速开始

### 环境要求

- **Python 3.8+**
- Windows 10 / 11，或 Android + [Termux](https://termux.dev/)

### 安装

```bash
git clone https://github.com/ouroooo/ToolsBox.git
cd ToolsBox
python -m pip install -r requirements.txt
```

<details>
<summary><b>Android / Termux 额外一步</b></summary>

若要读写手机共享存储（如 `/sdcard/Download`），需要先授权：

```bash
pkg install python
termux-setup-storage     # 弹窗里允许存储权限
```

</details>

### 运行

```bash
python main.py
```

启动后会先询问运行平台，直接回车即采用自动探测的结果：

```
运行平台
--------------------------------------------------------
  [1] Windows         使用 cls 清屏；路径形如 C:\Users\...
  [2] Android/Termux  使用 clear 清屏；可浏览 /sdcard 等常见目录
--------------------------------------------------------
请输入序号（默认 1）:
```

想跳过询问（脚本调用、桌面快捷方式）：

```bash
python main.py --platform=windows
python main.py --platform=android
```

> **Windows 用户请注意**：请在 cmd / PowerShell 里运行，**不要双击 `main.py`**。
> 双击时标准输入会立刻结束，程序会直接退出。

## 🧰 内置工具

### 1️⃣ xlsx 数据提取

从 xlsx 中提取「某列等于 / 包含某内容」的所有行，导出为新的 xlsx（表头原样保留）。

- 支持**多行表头**与合并单元格
- 匹配方式：精确 / 包含 / 不等于 / 为空 / 非空
- 导出前先预览命中行数与行号
- 输出：`output/成绩表_班级_15.xlsx`

### 2️⃣ 图片转 PDF

把一张或多张图片合并成一个 PDF。

- 支持 `jpg` `jpeg` `png`，另兼容 `bmp` `gif` `webp` `tif` `tiff`
- 三种尺寸模式：**原图尺寸** / **A4 自适应** / **固定宽度（mm）**
- 可增删图片、调整顺序、整目录批量添加
- 带透明通道的图片自动以白底填充，避免出现黑块

### 3️⃣ 多表行列提取

一次导入**多个 xlsx** 或**一个 xlsx 的多个工作表**，按条件提取后合并到一个新 xlsx。

- **行条件**：如「班级列包含 15」的数据行（可多条，AND 关系）
- **列条件**：如「列名包含 15」的整列（可多条）
- 两者同时使用则取交集
- 每个工作表输出为独立 sheet，重名自动加序号
- 输出：`output/原文件名_合并提取.xlsx`

### 4️⃣ PDF 单页提取

从 PDF 中取出指定的一页，另存为一个只含该页的新 PDF。

- 自动读取并显示原 PDF 的总页数，页码带范围校验
- 输出默认保存到 `output/原文件名_pageN.pdf`
- **Android 模式**下可浏览常见目录挑选 PDF

## 📁 目录结构

```
ToolsBox/
├── main.py                      # 程序入口：询问平台 → 工具选择菜单
├── requirements.txt
├── README.md                    # 本文件
├── README_ai.md                 # 面向 AI Agent / 开发者的技术手册
├── UNLICENSE                    # The Unlicense（公有领域）
├── core/                        # 公共基础设施
│   ├── runtime.py               # 运行平台：探测 / 询问 / 切换
│   ├── ui.py                    # 终端交互：标题、菜单、输入、提示
│   ├── paths.py                 # 路径：输出目录、常见目录、路径示例
│   └── registry.py              # 工具注册表
├── tools/                       # 各工具独立文件夹
│   ├── excel_extract/           # 【工具 1】xlsx 数据提取
│   ├── image_to_pdf/            # 【工具 2】图片转 PDF
│   ├── excel_multi_extract/     # 【工具 3】多表行列提取
│   └── pdf_page/                # 【工具 4】PDF 单页提取
└── output/                      # 默认输出目录（自动创建）
```

## ❓ 常见问题

<details>
<summary><b>Windows 下双击 <code>main.py</code> 窗口一闪而过？</b></summary>

这是预期行为。双击运行时标准输入立即结束（EOF），程序会把所有输入当作「取消」并退出。
请在 cmd 或 PowerShell 中执行 `python main.py`。

</details>

<details>
<summary><b>输出重定向到文件后中文乱码？</b></summary>

交互运行时中文正常；但重定向/管道输出时 Python 会按系统区域编码（简体中文 Windows 为 GBK）写出。
需要 UTF-8 时：

```bat
chcp 65001
set PYTHONIOENCODING=utf-8
python main.py
```

</details>

<details>
<summary><b>Termux 里选目录提示「目录不存在」？</b></summary>

先执行一次 `termux-setup-storage` 并授予存储权限，`/sdcard` 才会真正可访问。

</details>

<details>
<summary><b>提示找不到 <code>python</code> 命令？</b></summary>

Windows 下重装 Python 并勾选 **Add Python to PATH**，或改用 `py main.py`。

</details>

<details>
<summary><b>为什么每次启动都要问我平台？</b></summary>

因为两个平台的清屏命令和可用目录不同。默认值来自自动探测，直接回车即可；
也可以加 `--platform=windows` / `--platform=android` 永久跳过。

</details>

## 🧩 扩展：添加自己的工具

1. 在 `tools/` 下新建文件夹，例如 `tools/my_tool/`；
2. 在 `__init__.py` 中暴露入口：

   ```python
   NAME = "工具名"
   ORDER = 1                      # 数字越小越靠前，默认 99
   DESCRIPTION = "一句话说明"
   from .impl import run          # run() 为入口，自行完成交互
   ```

3. 在 `core/registry.py` 的 `TOOLS` 列表里登记该模块。

交互部分可以直接复用 `core/ui.py` 与 `core/paths.py`：

```python
from core import ui, paths

def run():
    ui.title("我的工具")
    name = ui.ask("请输入名称")
    if ui.ask_yes_no("确认", "y"):
        ...
    ui.success("完成")
    ui.pause()
```

> 更详细的实现约定（模块职责、依赖方向、读输入的两条铁律）见 [`README_ai.md`](README_ai.md)。

## 📄 许可证

本项目采用 [The Unlicense](UNLICENSE) —— **放弃全部著作权，贡献至公有领域**。

你可以自由地复制、修改、发布、使用、编译、销售或分发本软件，
无论是源代码形式还是编译后的二进制形式，无论用于商业还是非商业目的，
**无需署名，也无需保留任何版权声明**。

> 选择 Unlicense 而非传统软件许可证的原因：本项目主要由 AI 辅助生成
---

<div align="center">
<sub>如果这个工具箱帮到了你，欢迎点个 ⭐ Star</sub>
</div>
