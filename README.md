# Python工具箱 🛠️

<div align="center">

![Python工具箱](assets/icon.ico)

一个功能丰富、模块化设计的Python桌面工具箱，集成文件处理、PDF操作、图像处理、财务分析和AI功能于一体。

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

[功能特点](#-功能特点) • [安装指南](#-安装指南) • [快速开始](#-快速开始) • [模块开发](#-模块开发) • [项目结构](#-项目结构)

</div>

## ✨ 功能特点

### 📁 文件处理工具
- **Excel格式转换**: 批量将XLS格式转换为XLSX格式
- **列数据提取**: 从Excel文件中智能提取特定列数据
- **批量文件处理**: 批量重命名、移动或处理文件
- **智能数据提取**: 高级数据提取和分析功能

### 📄 PDF处理工具
- **PDF索引生成**: 自动生成PDF文档索引和目录
- **PDF合并**: 将多个PDF文件合并为一个文档

### 🖼️ 图像处理工具
- **证件照处理**: 自动生成标准尺寸证件照，支持背景去除
- **贴图制作**: 创建个性化贴图和标签

### 📊 财务分析工具（核心特色）
- **智能对账**: 自动化对账功能，支持多维度数据比对
- **审计雷达**: 财务风险识别和分析工具
- **对方科目分析**: 专业的财务科目分析功能

### 🔍 智能搜索工具
- **关键词搜索**: 全文关键词搜索和筛选
- **NLP文本聚类**: 基于自然语言处理的文本聚类分析
- **AI控制台**: 集成AI功能的交互式控制台

### 🤖 AI功能
- **背景移除**: 使用U2Net模型进行智能背景去除
- **文本向量化**: 集成text2vec-base-chinese模型进行中文文本处理
- **智能分析**: 基于机器学习的数据分析和预测

## 🚀 安装指南

### 环境要求
- Python 3.8+
- Windows 操作系统

### 安装步骤

1. **克隆仓库**
   ```bash
   git clone https://github.com/yourusername/python-toolbox.git
   cd python-toolbox
   ```

2. **创建虚拟环境**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

4. **运行程序**
   ```bash
   python main.py
   ```

### 打包为可执行文件

如果想将程序打包为独立的exe文件：

```bash
python build_exe.py
```

打包后的可执行文件将直接生成在项目根目录中（`Python工具箱.exe`）。打包过程会自动清理临时文件，不保留dist和build目录。

## 🎯 快速开始

启动程序后，您会看到一个现代化的界面，左侧是功能导航栏，右侧是功能区。

1. **选择功能**: 从左侧导航栏选择需要的功能模块
2. **设置参数**: 根据功能需求设置相关参数
3. **执行操作**: 点击相应按钮执行功能
4. **查看结果**: 在结果区域查看处理结果或下载输出文件

## 🧩 模块开发

Python工具箱采用"主程序 + 插件化模块"的架构，添加新功能非常简单，只需4步：

### 第一步：编写模块代码

使用以下标准提示词模板让AI为您生成代码：

> 我正在为一个 Python CustomTkinter GUI 项目添加新功能。
> 请帮我写一个 python 模块文件，保存在 modules 文件夹下。
>
> **功能需求：**
> (在这里写你想做什么，例如：帮我做一个 PDF 转 Word 的功能，要求能选择文件夹，批量转换)
>
> **代码结构要求（必须严格遵守）：**
> 1. **依赖库**：如果用到第三方库，请在代码最上方注释里告诉我需要 pip install 什么。
> 2. **核心逻辑与界面分离**：
>    - 写一个 `core_function(...)` 函数负责纯逻辑处理，接受参数，返回结果字符串，不要出现 print，用 log_callback 回调。
> 3. **界面类封装**：
>    - 定义一个类，例如 `class PdfToWordModule:`
>    - 在 `__init__` 中定义 `self.name = "PDF转Word"` (这是菜单显示的名字)。
>    - 必须包含 `render(self, parent_frame)` 方法。
>    - 在 `render` 方法的第一行，必须执行 `for widget in parent_frame.winfo_children(): widget.destroy()` 来清空页面。
>    - 所有的 UI 组件（按钮、输入框）都使用 `customtkinter`，并挂载在 `parent_frame` 上。
>    - 耗时操作（如转换文件）必须使用 `threading` 开启新线程，不要卡死界面。

### 第二步：放入文件

1. 在项目的 `modules` 文件夹里新建一个 `.py` 文件（例如 `pdf_tool.py`）
2. 把 AI 生成的代码粘贴进去

### 第三步：注册功能

打开根目录的 `main.py`，在两处添加代码：

```python
# 1. 在顶部导入 (约第 5-8 行)
from modules.xls_to_xlsx import XLSToXLSXModule
from modules.column_extractor import ColumnExtractorModule
from modules.pdf_tool import PdfToWordModule  # <--- 新增这行

# ... 中间代码不变 ...

# 2. 在 App 类的 __init__ 中注册 (约第 48 行)
self.modules = [
    XLSToXLSXModule(),
    ColumnExtractorModule(),
    PdfToWordModule(),  # <--- 新增这行
]
```

### 第四步：运行测试

重新启动程序，您会看到新功能已经出现在左侧导航栏中！

## 📁 项目结构

```
Python工具箱/
│
├── 主程序和配置文件
│   ├── main.py                 # 主程序入口，负责界面框架和功能集成
│   ├── build_exe.py            # 打包脚本，用于生成可执行文件
│   ├── Python工具箱.spec       # PyInstaller打包配置文件
│   ├── requirements.txt        # 项目依赖包列表
│   └── README.md               # 项目说明文档
│
├── 资源目录
│   └── assets/
│       ├── fonts/
│       │   └── simsun.ttc      # 中文字体文件（宋体）
│       ├── models/
│       │   └── u2net.onnx      # AI抠图模型文件
│       └── nlp/                 # 自然语言处理模型目录
│           └── text2vec-base-chinese/ # 中文文本向量化模型（魔塔社区开源模型）
│               ├── model.safetensors # 安全张量模型文件（390MB）
│               ├── pytorch_model.bin # PyTorch模型文件（390MB）
│               ├── onnx/       # ONNX模型格式目录（多种优化版本）
│               ├── openvino/   # OpenVINO模型格式目录
│               └── ...         # 其他模型配置和词汇表文件（共30+文件）
│
├── 功能模块目录
│   └── modules/
│       ├── xls_to_xlsx.py      # XLS转XLSX功能模块
│       ├── column_extractor.py # 列数据提取功能模块
│       ├── file_batch_tool.py  # 文件批量处理功能模块
│       ├── pdf_indexer.py      # PDF索引功能模块
│       ├── pdf_merger.py       # PDF合并功能模块
│       ├── id_photo_tool.py    # 证件照处理功能模块
│       ├── ai_console.py       # AI控制台功能模块
│       ├── keyword_search.py   # 关键词搜索功能模块
│       ├── smart_extractor.py  # 智能提取功能模块
│       ├── audit_radar_module.py # 审计雷达总接口模块
│       └── audit_radar/        # 审计雷达核心算法包
│           ├── model.py        # PyTorch神经网络结构
│           ├── data_processor.py # 数据清洗逻辑
│           └── engine.py       # 训练和推理逻辑
│
├── 用户数据目录
│   └── user_data/
│       ├── ai_config.json      # AI配置文件
│       └── schemas/            # 数据模式定义
│
├── 日志目录
│   └── log/                    # 程序运行日志目录
│
├── 构建文件
│   └── Python工具箱.exe         # 打包后的可执行文件（直接生成在根目录）
│
└── 工具配置目录
    └── .codebuddy/
        └── rules/
            └── 禁止修改代码.mdc  # 代码修改限制配置
```

## 🛠️ 技术栈

- **界面框架**: [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - 现代化的Tkinter界面库
- **Excel处理**: openpyxl, xlrd, pandas - Excel文件读写和数据处理
- **PDF处理**: PyMuPDF, pdfplumber - PDF文件处理
- **图像处理**: Pillow, rembg - 图像处理库和背景移除
- **AI模型**: sentence-transformers, PyTorch, ONNX Runtime - 自然语言处理和模型推理
  - 注：NLP模型text2vec-base-chinese来自[魔塔社区](https://modelscope.cn)开源项目
- **数据分析**: scikit-learn, scipy - 数据分析和机器学习
- **打包工具**: PyInstaller - 将Python程序打包为可执行文件
