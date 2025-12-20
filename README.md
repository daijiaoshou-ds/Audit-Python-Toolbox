# 🐍 Python 审计工具箱 (Audit Python Toolbox)

<div align="center">

![Icon](assets/icon.ico)

专为财务/审计人员打造的 Python 自动化工具箱。  
集成 Excel 批量处理、OCR 智能提取、NLP 语义聚类、PDF 操作等核心功能。  
**无需安装 Python，开箱即用（Release版）。**

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

[功能特性](#-功能特性) • [安装指南](#-安装指南) • [模型下载](#-模型下载-重要) • [开发文档](#-开发文档)

</div>

## ✨ 功能特性

### 📊 审计与财务分析 (Core)
- **🧠 智能语义聚类 (Pro)**: 使用 `text2vec` 深度学习模型，自动分析数万条摘要，归纳业务类型（支持科目分割、金额过滤）。
- **🔍 智能对账**: 自动化多维度数据比对。
- **📡 审计雷达**: 财务风险识别与异常检测。

### 📁 办公自动化
- **Excel 工具**: 格式互转 (xls/xlsx/xlsm)、指定列提取、批量重命名。
- **PDF 工具**: 自动生成索引、合并/分拆、Word/图片转PDF。
- **图像工具**: 证件照换底/改尺寸（AI 自动抠图）、贴纸制作。

---

## 🚀 安装指南

### 方式一：直接使用 (推荐给财务同事)
1.  前往 [Releases 页面](https://github.com/daijiaoshou-ds/Audit-Python-Toolbox/releases) 下载最新版压缩包。
2.  解压后，**直接双击 `Python工具箱.exe`** 即可使用。
    *   *注意：请确保 `assets` 文件夹与 exe 在同一目录下。*

### 方式二：源码运行 (推荐给开发者)
1.  **克隆仓库**:
    ```bash
    git clone https://github.com/daijiaoshou-ds/Audit-Python-Toolbox.git
    ```
2.  **安装依赖**:
    ```bash
    pip install -r requirements.txt
    ```
3.  **下载模型** (见下文)。
4.  **运行**: `python main.py`

---

## 📥 模型下载 (重要!)

为了减小体积，GitHub 仓库不包含大模型文件。源码运行需手动下载：

1.  **NLP 模型 (text2vec-base-chinese)**:
    *   下载地址: [HuggingFace](https://huggingface.co/shibing624/text2vec-base-chinese) 或 [魔搭社区](https://modelscope.cn/models/shibing624/text2vec-base-chinese)
    *   存放位置: `assets/models/nlp/text2vec-base-chinese/`
2.  **抠图模型 (u2net)**:
    *   程序首次运行会自动下载，或手动下载 `u2net.onnx` 放入 `assets/models/`。

---

## 🛠️ 技术栈
- **GUI**: CustomTkinter
- **AI Core**: PyTorch, Sentence-Transformers, ONNX Runtime
- **Data**: Pandas, Scikit-learn
- **Build**: PyInstaller (自动化构建脚本)