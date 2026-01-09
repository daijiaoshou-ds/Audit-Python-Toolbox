# 😸 基米工具箱 (Hajimi Toolbox)

<div align="center">

<img src="assets/icon.ico" width="128" height="128" alt="Logo">

**专为财务/审计人员打造的 AI 自动化工作台**  
集成 OCR 智能提取、NLP 语义分析、审计雷达、批量办公处理等核心功能。  
**懒加载架构 · 极速启动 · 离线可用**

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![Build](https://img.shields.io/badge/Build-Inno_Setup-orange.svg)](https://jrsoftware.org/isinfo.php)
[![Platform](https://img.shields.io/badge/Platform-Windows_10%2F11-lightgrey.svg)](https://www.microsoft.com/windows)

[核心功能](#-核心功能) • [下载与运行](#-下载与运行) • [模型配置](#-ai-模型配置) • [技术栈](#-技术栈)

</div>

---

## ✨ 核心功能

### 🧠 AI 智能审计 (Core)
- **OCR 智能文档提取**: "双引擎"架构（Eye-AI + Brain-AI），支持 PDF/图片转结构化 Excel，自定义提取字段。
- **NLP 语义聚类**: 基于 `text2vec` 深度学习模型，自动分析数万条摘要，智能归纳业务类型。
- **审计雷达**: 财务异常检测，自动扫描会计分录风险点。
- **银行流水核查**: 智能匹配算法，解决金额/日期微小差异的对账难题。
- **对方科目分析**: `穷举计算`+`奥卡姆剃刀算法`+`记忆得分EMA算法`+`记忆指纹泛化`一套组合拳👊，解决多借多贷分析问题。

### ⚡ 办公自动化 (Efficiency)
- **🛡️ 全局中断机制**: 任何耗时任务（如批量处理、AI运算）均支持**一键立即终止**，安全不坏档。
- **Excel 工具**: 格式互转 (xls/xlsx/xlsm)、多表指定列提取（支持Polars）、数据清洗。
- **PDF 工具**: 自动生成页码索引、合并/分拆/格式转换（支持 Word/Excel 转 PDF）。
- **图像工具**: 证件照一键换底/改尺寸（本地 AI 抠图）、表情包制作。

---

## 🚀 下载与运行

### 方式一：安装包 (推荐)
最简单的使用方式，无需配置 Python 环境。

1.  前往 [Releases 页面](#) 下载最新版 **`Hajimi_Setup_v0.3.exe`**。
2.  双击安装 (支持自定义路径，智能记忆上次安装位置)。
3.  桌面双击 **基米工具箱** 图标即可秒开使用。

### 方式二：源码运行
适合想要二次开发或查看源码的用户。

1.  **克隆仓库**:
    ```bash
    git clone https://github.com/daijiaoshou-ds/Audit-Python-Toolbox.git
    ```
2.  **安装依赖**:
    ```bash
    pip install -r requirements.txt
    ```
3.  **运行程序**:
    ```bash
    python main.py
    ```

---

## 📥 AI 模型配置

本工具采用 **本地 + 在线** 混合模式，请根据需求配置：

### 1. 本地离线模型 (必需)
为了减小体积，部分大模型需手动下载放入 `assets/models` 目录：

*   **NLP 语义模型 (`text2vec`)**:
    *   下载地址: [HuggingFace](https://huggingface.co/shibing624/text2vec-base-chinese)
    *   存放位置: `assets/models/nlp/text2vec-base-chinese/`
*   **AI 抠图模型 (`u2net`)**:
    *   程序首次运行证件照功能时会自动下载。
    *   或手动下载 `u2net.onnx` 放入 `assets/models/` 根目录。

### 2. 在线大模型 (可选)
用于“智能文档提取”和“AI 控制台”功能：
*   在软件左侧点击 **"AI 控制台"**。
*   配置你的 API Key (支持 OpenAI / Moonshot / DeepSeek 等)。

---

## 🛠️ 技术栈

*   **UI 框架**: CustomTkinter (高 DPI 适配 / Win32 API 高清渲染)
*   **AI 引擎**: PyTorch, Sentence-Transformers, ONNX Runtime
*   **数据处理**: Pandas, Polars, Scipy
*   **架构设计**: 模块化懒加载 (Lazy Loading) + 视图缓存
*   **打包发布**: PyInstaller (Hidden-Import 注入) + Inno Setup (LZMA2 极限压缩)

---

<div align="center">
Made with ❤️ by Daijiaoshou
</div>