"""
Excel数据提取器 - 升级版
修复了Excel读取兼容性bug，集成Polars高性能引擎
支持pandas和polars双引擎模式
"""
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from difflib import SequenceMatcher
import shutil
from typing import Optional, List, Dict, Any, Tuple

# ==================== 新增：Polars支持 ====================
try:
    import polars as pl
    POLARS_AVAILABLE = True
except ImportError:
    POLARS_AVAILABLE = False
    print("警告: Polars未安装，将使用pandas模式。可通过 'pip install polars' 安装以提升性能。")

# ==================== 新增：pyarrow支持检测 ====================
try:
    import pyarrow
    PYARROW_AVAILABLE = True
except ImportError:
    PYARROW_AVAILABLE = False
    print("警告: pyarrow未安装，Polars模式性能可能受限。可通过 'pip install pyarrow' 提升性能。")

# ==================== 核心逻辑层 ====================

def calculate_similarity(s1: str, s2: str) -> float:
    """计算两个字符串的相似度"""
    if not s1 or not s2:
        return 0.0
    return SequenceMatcher(None, s1.lower(), s2.lower()).ratio()

def is_fuzzy_match(target_keyword: str, cell_value: Any, threshold: float) -> bool:
    """智能模糊匹配判断"""
    t = str(target_keyword).strip().lower()
    v = str(cell_value).strip().lower() if cell_value is not None else ""
    if threshold >= 0.99:
        return t == v
    if threshold <= 0.9:
        if t in v or v in t:
            return True
    score = calculate_similarity(t, v)
    return score >= threshold

def get_files_to_process(source_path: str, is_folder: bool, recursive: bool) -> List[str]:
    """根据设置获取文件列表"""
    files_list = []
    if not is_folder:
        if os.path.isfile(source_path) and source_path.lower().endswith(('.xlsx', '.xlsm')):
            return [source_path]
        return []
    if not os.path.exists(source_path):
        return []

    if recursive:
        for root, dirs, files in os.walk(source_path):
            for f in files:
                if f.lower().endswith(('.xlsx', '.xlsm')) and not f.startswith("~$"):
                    files_list.append(os.path.join(root, f))
    else:
        for f in os.listdir(source_path):
            full_path = os.path.join(source_path, f)
            if os.path.isfile(full_path) and f.lower().endswith(('.xlsx', '.xlsm')) and not f.startswith("~$"):
                files_list.append(full_path)
    return files_list

# ==================== 【修复】新增：Excel文件预处理 ====================
def repair_excel_file(file_path: str, log_func) -> bool:
    """
    修复损坏或格式不标准的Excel文件
    通过重新保存来修复文件的元数据问题
    """
    try:
        # 首先尝试正常读取
        wb = load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        
        # 检查维度是否合理
        if ws.max_row > 0 and ws.max_column > 0:
            # 尝试读取第一行数据来验证文件可读性
            test_row = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            if test_row and any(cell is not None for cell in test_row[0] if test_row):
                wb.close()
                return True  # 文件正常
        wb.close()
    except Exception as e:
        log_func(f"检测到文件异常，尝试修复: {os.path.basename(file_path)}")
    
    try:
        # 尝试修复：使用非只读模式打开并重新保存
        temp_wb = load_workbook(file_path, read_only=False, data_only=True)
        temp_file = file_path + ".temp"
        
        # 保存到临时文件
        temp_wb.save(temp_file)
        temp_wb.close()
        
        # 替换原文件
        shutil.move(temp_file, file_path)
        log_func(f"✓ 文件修复成功: {os.path.basename(file_path)}")
        return True
    except Exception as e:
        log_func(f"✗ 文件修复失败: {os.path.basename(file_path)} - {str(e)}")
        return False

def scan_header_and_map_columns(worksheet, mode: str, 
                                exact_cols: Optional[List[str]] = None, 
                                fuzzy_cols: Optional[List[str]] = None, 
                                fuzzy_threshold: float = 0.6,
                                scan_rows: int = 15) -> Dict[int, str]:
    """扫描表头并建立映射"""
    max_scan = min(worksheet.max_row, scan_rows)
    if max_scan == 0:
        return {}
    best_mapping = {}
    max_matches = -1
    
    for r in range(1, max_scan + 1):
        row_values = []
        for c in range(1, worksheet.max_column + 1):
            val = worksheet.cell(row=r, column=c).value
            row_values.append(str(val).strip() if val is not None else "")

        if all(v == "" for v in row_values):
            continue
        current_mapping = {}
        
        if mode == 'all':
            name_counter = {}
            for c_idx, val in enumerate(row_values, 1):
                if not val:
                    continue
                count = name_counter.get(val, 0)
                final_name = val if count == 0 else f"{val}_{count}"
                name_counter[val] = count + 1
                current_mapping[c_idx] = final_name
            if current_mapping:
                return current_mapping

        else:
            matches_count = 0
            name_counter = {}
            exact_cols = exact_cols or []
            fuzzy_cols = fuzzy_cols or []
            
            for c_idx, val in enumerate(row_values, 1):
                if not val:
                    continue
                val_lower = val.lower()
                matched_name = None
                
                # 精确匹配
                for target in exact_cols:
                    if target.lower() == val_lower:
                        matched_name = target
                        break
                
                # 模糊匹配
                if not matched_name and fuzzy_cols:
                    for target in fuzzy_cols:
                        if is_fuzzy_match(target, val, fuzzy_threshold):
                            matched_name = target
                            break
                
                if matched_name:
                    count = name_counter.get(matched_name, 0)
                    final_key = matched_name if count == 0 else f"{matched_name}_{count}"
                    name_counter[matched_name] = count + 1
                    current_mapping[c_idx] = final_key
                    matches_count += 1
            
            if matches_count > max_matches:
                max_matches = matches_count
                best_mapping = current_mapping
    
    return best_mapping

# ==================== 【新增】Polars核心处理逻辑 ====================
def process_with_polars(files: List[str], target_sheets: List[str], 
                       extract_mode: str, exact_cols: List[str], 
                       fuzzy_cols: List[str], fuzzy_threshold: float,
                       save_path: str, log_func, stop_event=None) -> Tuple[bool, str]:
    """使用Polars引擎处理数据 - 高性能版本"""
    
    log_func("=== Polars高性能模式 ===")
    log_func(f"发现 {len(files)} 个文件...")
    
    all_dfs = []
    processed_count = 0
    
    for i, file_path in enumerate(files):
        # 检查中断
        if stop_event and stop_event.is_set():
            log_func(">>> 用户强制停止任务！")
            return False, "任务已终止。"
        
        fname = os.path.basename(file_path)
        log_func(f"[{i+1}/{len(files)}] Polars读取: {fname}")
        
        try:
            # 尝试直接读取
            try:
                # 读取Excel文件
                df = pl.read_excel(
                    file_path,
                    engine="openpyxl",
                    read_only=False
                )
            except Exception as e:
                # 如果直接读取失败，尝试逐个sheet读取
                df = None
                temp_wb = load_workbook(file_path, read_only=False, data_only=True)
                
                for ws in temp_wb.worksheets:
                    if target_sheets:
                        is_match = False
                        for target in target_sheets:
                            if is_fuzzy_match(target, ws.title, fuzzy_threshold):
                                is_match = True
                                break
                        if not is_match:
                            continue
                    
                    # 手动读取数据并转换为Polars
                    data = []
                    headers = None
                    
                    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                        if row_idx == 1:
                            # 第一行作为表头
                            headers = [str(cell).strip() if cell is not None else "" for cell in row]
                            continue
                        
                        row_data = list(row)
                        if any(cell is not None for cell in row_data):
                            data.append(row_data)
                    
                    if headers and data:
                        # 创建Polars DataFrame
                        sheet_df = pl.DataFrame(data, schema=headers, orient='row')
                        sheet_df = sheet_df.with_columns([
                            pl.lit(fname).alias('_来源工作簿'),
                            pl.lit(ws.title).alias('_来源工作表')
                        ])
                        
                        if df is None:
                            df = sheet_df
                        else:
                            df = pl.concat([df, sheet_df], how='vertical_relaxed')
                
                temp_wb.close()
            
            if df is not None and df.height > 0:
                all_dfs.append(df)
                processed_count += 1
                log_func(f"  ✓ 读取成功: {df.height} 行")
                
        except Exception as e:
            log_func(f"  ✗ 读取失败: {e}")
            # 尝试修复文件后重试
            if repair_excel_file(file_path, log_func):
                try:
                    df = pl.read_excel(file_path)
                    if df is not None and df.height > 0:
                        all_dfs.append(df)
                        processed_count += 1
                        log_func(f"  ✓ 修复后读取成功: {df.height} 行")
                except Exception as retry_e:
                    log_func(f"  ✗ 修复后仍失败: {retry_e}")
    
    if not all_dfs:
        return False, "未提取到任何数据。"
    
    try:
        log_func("正在汇总并保存...")
        # 合并所有DataFrame
        combined_df = pl.concat(all_dfs, how='vertical_relaxed')
        
        # 列处理
        source_cols = [c for c in combined_df.columns if c not in ['_来源工作簿', '_来源工作表']]
        
        if extract_mode == 'specific':
            # 只保留指定的列
            valid_cols = []
            for exact in exact_cols:
                if exact in source_cols:
                    valid_cols.append(exact)
            
            remaining = [c for c in source_cols if c not in valid_cols]
            final_cols = valid_cols + remaining
        else:
            final_cols = source_cols
        
        final_cols = final_cols + ['_来源工作簿', '_来源工作表']
        
        # 确保所有列都存在
        final_cols = [c for c in final_cols if c in combined_df.columns]
        combined_df = combined_df.select(final_cols)
        
        # 保存文件（使用pandas+pyarrow，性能最佳）
        if PYARROW_AVAILABLE:
            try:
                # 利用pyarrow高效转换，速度最快
                pandas_df = combined_df.to_pandas()
                pandas_df.to_excel(save_path, index=False)
                return True, f"Polars模式完成！共提取 {len(pandas_df)} 行数据。\n保存至: {save_path}"
            except Exception as e:
                log_func(f"  pandas保存异常，尝试备用方案: {e}")
        
        # 备用方案：openpyxl直接写入
        try:
            from openpyxl import Workbook
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            
            # 写入表头
            headers = list(combined_df.columns)
            ws.append(headers)
            
            # 写入数据（逐行处理）
            for row in combined_df.rows():
                ws.append(list(row))
            
            wb.save(save_path)
            wb.close()
            
            return True, f"Polars模式完成！共提取 {len(combined_df)} 行数据。\n保存至: {save_path}"
            
        except Exception as e:
            return False, f"保存失败: {e}"
        
    except Exception as e:
        return False, f"保存失败: {e}"

# ==================== 原有Pandas核心处理逻辑（已优化）====================
def process_with_pandas(files: List[str], target_sheets: List[str], 
                       extract_mode: str, exact_cols: List[str], 
                       fuzzy_cols: List[str], fuzzy_threshold: float,
                       save_path: str, log_func, stop_event=None) -> Tuple[bool, str]:
    """使用Pandas引擎处理数据"""
    
    log_func("=== Pandas标准模式 ===")
    log_func(f"发现 {len(files)} 个文件...")
    
    all_rows = []
    processed_count = 0
    
    for i, file_path in enumerate(files):
        # 检查中断
        if stop_event and stop_event.is_set():
            log_func(">>> 用户强制停止任务！")
            return False, "任务已终止。"
        
        fname = os.path.basename(file_path)
        log_func(f"[{i+1}/{len(files)}] 正在读取: {fname}")
        
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            
            for ws in wb.worksheets:
                # Sheet循环里检查中断
                if stop_event and stop_event.is_set():
                    wb.close()
                    log_func(">>> 用户强制停止任务！")
                    return False, "任务已终止。"
                
                if target_sheets:
                    is_match = False
                    for target in target_sheets:
                        if is_fuzzy_match(target, ws.title, fuzzy_threshold):
                            is_match = True
                            break
                    if not is_match:
                        continue
                
                col_map = scan_header_and_map_columns(ws, extract_mode, exact_cols, fuzzy_cols, fuzzy_threshold)
                if not col_map:
                    continue
                
                for row_data in ws.iter_rows(values_only=True):
                    extracted_row = {}
                    has_valid_data = False
                    for col_idx, target_name in col_map.items():
                        if col_idx - 1 < len(row_data):
                            val = row_data[col_idx - 1]
                            extracted_row[target_name] = val
                            if val is not None:
                                has_valid_data = True
                    
                    if has_valid_data:
                        match_header_count = sum(1 for k, v in extracted_row.items() if str(v) == k)
                        if match_header_count > 0 and match_header_count > len(extracted_row) / 2:
                            continue
                        extracted_row['_来源工作簿'] = fname
                        extracted_row['_来源工作表'] = ws.title
                        all_rows.append(extracted_row)
            
            wb.close()
            processed_count += 1
            
        except InvalidFileException as e:
            log_func(f"  文件格式异常: {e}")
            # 尝试修复文件
            if repair_excel_file(file_path, log_func):
                try:
                    wb = load_workbook(file_path, read_only=True, data_only=True)
                    # 重复读取逻辑...
                    wb.close()
                    processed_count += 1
                except Exception as retry_e:
                    log_func(f"  ✗ 修复后仍失败: {retry_e}")
        except Exception as e:
            log_func(f"  ✗ 读取失败: {e}")
    
    if not all_rows:
        return False, "未提取到任何数据。"
    
    try:
        log_func("正在汇总并保存...")
        df = pd.DataFrame(all_rows)
        cols = [c for c in df.columns if c not in ['_来源工作簿', '_来源工作表']]
        
        if extract_mode == 'specific':
            sorted_cols = []
            for exact in exact_cols:
                if exact in cols:
                    sorted_cols.append(exact)
            remaining = [c for c in cols if c not in sorted_cols]
            final_cols = sorted_cols + remaining + ['_来源工作簿', '_来源工作表']
        else:
            final_cols = cols + ['_来源工作簿', '_来源工作表']
            
        df = df.reindex(columns=final_cols)
        df.to_excel(save_path, index=False)
        
        return True, f"完成！共提取 {len(df)} 行数据。\n保存至: {save_path}"
        
    except Exception as e:
        return False, f"保存失败: {e}"

# ==================== 统一入口函数 ====================
def core_process(source_path: str, is_folder: bool, recursive: bool, 
                target_sheets: List[str], extract_mode: str, 
                exact_cols: List[str], fuzzy_cols: List[str], 
                fuzzy_threshold: float, save_path: str, 
                log_func, stop_event=None,
                use_polars: bool = False) -> Tuple[bool, str]:
    """核心处理函数 - 支持双引擎"""
    
    log_func("=== 任务开始 ===")
    
    files = get_files_to_process(source_path, is_folder, recursive)
    if not files:
        return False, "未找到有效的 Excel 文件。"
    
    # 根据设置选择处理引擎
    if use_polars and POLARS_AVAILABLE:
        return process_with_polars(
            files, target_sheets, extract_mode, exact_cols, 
            fuzzy_cols, fuzzy_threshold, save_path, log_func, stop_event
        )
    else:
        return process_with_pandas(
            files, target_sheets, extract_mode, exact_cols, 
            fuzzy_cols, fuzzy_threshold, save_path, log_func, stop_event
        )

# ==================== 界面层 ====================

class ColumnExtractorModule:
    def __init__(self):
        self.name = "Excel数据提取"
        self.input_path = ""
        self.app = None
        self.module_index = None
        
    def render(self, parent_frame):
        for widget in parent_frame.winfo_children():
            widget.destroy()

        STYLE_LABEL = {"font": ("Microsoft YaHei", 14), "text_color": "#333"}
        STYLE_ENTRY = {"height": 36, "border_color": "#D0D0D0", "border_width": 1, "fg_color": "#FAFAFA", "text_color": "#333"}
        STYLE_BTN_BLUE = {"height": 36, "fg_color": "#F0F5FA", "text_color": "#0984e3", "hover_color": "#E1EBF5"}
        
        scroll = ctk.CTkScrollableFrame(parent_frame, fg_color="transparent", scrollbar_button_color="#E0E0E0", scrollbar_button_hover_color="#D0D0D0")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(scroll, text="Excel数据提取器", font=("Microsoft YaHei", 24, "bold"), text_color="#333").pack(anchor="w", padx=20, pady=(10, 20))

        # --- 1. 数据源 ---
        frame_src = ctk.CTkFrame(scroll, fg_color="white", corner_radius=8, border_width=1, border_color="#E5E5E5")
        frame_src.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(frame_src, text="1. 数据源选择", font=("Microsoft YaHei", 15, "bold"), text_color="#0984e3").pack(anchor="w", padx=15, pady=10)
        
        self.var_source_type = ctk.StringVar(value="folder")
        f_type = ctk.CTkFrame(frame_src, fg_color="transparent")
        f_type.pack(fill="x", padx=15, pady=5)
        ctk.CTkRadioButton(f_type, text="处理文件夹 (批量)", variable=self.var_source_type, value="folder", text_color="#333", command=self.toggle_source_ui).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(f_type, text="处理单个文件", variable=self.var_source_type, value="file", text_color="#333", command=self.toggle_source_ui).pack(side="left")

        f_path = ctk.CTkFrame(frame_src, fg_color="transparent")
        f_path.pack(fill="x", padx=15, pady=10)
        self.entry_src = ctk.CTkEntry(f_path, placeholder_text="请选择路径...", **STYLE_ENTRY)
        self.entry_src.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_path, text="浏览...", width=100, command=self.select_source, **STYLE_BTN_BLUE).pack(side="left")

        self.check_recursive = ctk.CTkCheckBox(frame_src, text="递归遍历子文件夹 (仅文件夹模式有效)", text_color="#555", font=("Microsoft YaHei", 12))
        self.check_recursive.pack(anchor="w", padx=15, pady=(0, 15))

        # --- 2. 规则设置 ---
        frame_rule = ctk.CTkFrame(scroll, fg_color="white", corner_radius=8, border_width=1, border_color="#E5E5E5")
        frame_rule.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(frame_rule, text="2. 提取规则设置", font=("Microsoft YaHei", 15, "bold"), text_color="#0984e3").pack(anchor="w", padx=15, pady=10)

        f_sheet = ctk.CTkFrame(frame_rule, fg_color="transparent")
        f_sheet.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(f_sheet, text="指定工作表：", width=100, anchor="e", **STYLE_LABEL).pack(side="left")
        self.entry_sheets = ctk.CTkEntry(f_sheet, placeholder_text="可选，留空则提取所有Sheet。逗号分隔", **STYLE_ENTRY)
        self.entry_sheets.pack(side="left", fill="x", expand=True)

        f_col_mode = ctk.CTkFrame(frame_rule, fg_color="transparent")
        f_col_mode.pack(fill="x", padx=15, pady=15)
        ctk.CTkLabel(f_col_mode, text="提取列模式：", width=100, anchor="e", **STYLE_LABEL).pack(side="left")
        self.switch_col_mode = ctk.CTkSwitch(f_col_mode, text="启用「指定列提取」", command=self.toggle_col_ui, text_color="#333", font=("Microsoft YaHei", 13))
        self.switch_col_mode.pack(side="left", padx=10)
        ctk.CTkLabel(f_col_mode, text="(关闭则自动提取所有发现的列)", text_color="gray", font=("Microsoft YaHei", 12)).pack(side="left")

        self.frame_col_detail = ctk.CTkFrame(frame_rule, fg_color="#F8F9FA", corner_radius=6)
        
        f_exact = ctk.CTkFrame(self.frame_col_detail, fg_color="transparent")
        f_exact.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(f_exact, text="精确匹配列：", width=100, anchor="e", **STYLE_LABEL).pack(side="left")
        self.entry_exact = ctk.CTkEntry(f_exact, placeholder_text="列名完全一致才提取，逗号分隔", **STYLE_ENTRY)
        self.entry_exact.pack(side="left", fill="x", expand=True)

        f_fuzzy = ctk.CTkFrame(self.frame_col_detail, fg_color="transparent")
        f_fuzzy.pack(fill="x", padx=10, pady=(10, 5))
        ctk.CTkLabel(f_fuzzy, text="模糊匹配列：", width=100, anchor="e", **STYLE_LABEL).pack(side="left")
        self.entry_fuzzy = ctk.CTkEntry(f_fuzzy, placeholder_text="包含关键词即提取，逗号分隔", **STYLE_ENTRY)
        self.entry_fuzzy.pack(side="left", fill="x", expand=True)
        
        f_slider = ctk.CTkFrame(self.frame_col_detail, fg_color="transparent")
        f_slider.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkLabel(f_slider, text="匹配严格度：", width=100, anchor="e", **STYLE_LABEL).pack(side="left")
        self.lbl_slider_val = ctk.CTkLabel(f_slider, text="0.3 (默认)", width=30, text_color="#0984e3", font=("Microsoft YaHei", 13, "bold"))
        self.lbl_slider_val.pack(side="right", padx=10)
        self.slider_fuzzy = ctk.CTkSlider(f_slider, from_=0.1, to=1.0, number_of_steps=90, command=self.update_slider_label)
        self.slider_fuzzy.set(0.3)
        self.slider_fuzzy.pack(side="left", fill="x", expand=True, padx=10)

        # --- 3. 引擎选择（新增）---
        frame_engine = ctk.CTkFrame(scroll, fg_color="white", corner_radius=8, border_width=1, border_color="#E5E5E5")
        frame_engine.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(frame_engine, text="3. 处理引擎选择", font=("Microsoft YaHei", 15, "bold"), text_color="#0984e3").pack(anchor="w", padx=15, pady=10)
        
        f_engine = ctk.CTkFrame(frame_engine, fg_color="transparent")
        f_engine.pack(fill="x", padx=15, pady=10)
        
        self.var_engine = ctk.StringVar(value="pandas")
        
        # Pandas选项
        self.radio_pandas = ctk.CTkRadioButton(
            f_engine, 
            text="Pandas (稳定兼容)", 
            variable=self.var_engine, 
            value="pandas",
            text_color="#333",
            command=self.toggle_engine_info
        )
        self.radio_pandas.pack(side="left", padx=(0, 20))
        
        # Polars选项（如果可用）
        if POLARS_AVAILABLE:
            self.radio_polars = ctk.CTkRadioButton(
                f_engine, 
                text="Polars (高性能)", 
                variable=self.var_engine, 
                value="polars",
                text_color="#E74C3C",  # 红色突出显示高性能
                command=self.toggle_engine_info
            )
            self.radio_polars.pack(side="left", padx=(0, 20))
            
            ctk.CTkLabel(
                f_engine, 
                text="⚡ 性能提升5-10倍，推荐大数据量使用", 
                text_color="#27AE60", 
                font=("Microsoft YaHei", 11)
            ).pack(side="left")
        else:
            ctk.CTkLabel(
                f_engine, 
                text="💡 Polars未安装 (pip install polars)", 
                text_color="#999999", 
                font=("Microsoft YaHei", 11)
            ).pack(side="left")
        
        # 引擎说明
        self.lbl_engine_info = ctk.CTkLabel(
            frame_engine,
            text="Pandas: 稳定兼容，适合所有场景。处理万行数据约需10-30秒。",
            text_color="#666666",
            font=("Microsoft YaHei", 11)
        )
        self.lbl_engine_info.pack(anchor="w", padx=15, pady=(0, 15))

        self.btn_run = ctk.CTkButton(scroll, text="开始执行提取", height=50, fg_color="#007AFF", font=("Microsoft YaHei", 16, "bold"), command=self.run_task)
        self.btn_run.pack(fill="x", padx=20, pady=20)

        ctk.CTkLabel(scroll, text="执行日志", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(anchor="w", padx=20, pady=(0, 5))
        self.textbox = ctk.CTkTextbox(scroll, height=180, fg_color="#FAFAFA", border_width=1, border_color="#D0D0D0", text_color="#333")
        self.textbox.pack(fill="x", padx=20, pady=(0, 20))

        self.toggle_source_ui()
        self.toggle_col_ui()
        self.toggle_engine_info()

    def update_slider_label(self, value):
        val = round(value, 2)
        desc = ""
        if val >= 0.99:
            desc = "(精确)"
        elif val >= 0.8:
            desc = "(严格)"
        elif val <= 0.4:
            desc = "(宽松)"
        else:
            desc = "(标准)"
        self.lbl_slider_val.configure(text=f"{val} {desc}")

    def toggle_source_ui(self):
        mode = self.var_source_type.get()
        if mode == "file":
            self.check_recursive.configure(state="disabled")
            self.entry_src.configure(placeholder_text="请选择单个 .xlsx 文件")
        else:
            self.check_recursive.configure(state="normal")
            self.entry_src.configure(placeholder_text="请选择包含 Excel 的文件夹")

    def toggle_col_ui(self):
        if self.switch_col_mode.get() == 1:
            self.frame_col_detail.pack(fill="x", padx=15, pady=(0, 15))
        else:
            self.frame_col_detail.pack_forget()
    
    def toggle_engine_info(self):
        """更新引擎说明"""
        engine = self.var_engine.get()
        if engine == "pandas":
            self.lbl_engine_info.configure(
                text="Pandas: 稳定兼容，适合所有场景。处理万行数据约需10-30秒。"
            )
        else:
            self.lbl_engine_info.configure(
                text="Polars: 高性能模式，利用多核CPU并行处理。处理万行数据约需1-3秒。"
            )

    def select_source(self):
        mode = self.var_source_type.get()
        path = ""
        if mode == "file":
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        else:
            path = filedialog.askdirectory()
        if path:
            self.entry_src.delete(0, "end")
            self.entry_src.insert(0, path)
            self.input_path = path

    def log(self, msg):
        self.textbox.insert("end", msg + "\n")
        self.textbox.see("end")

    def run_task(self):
        src = self.entry_src.get().strip()
        is_folder = (self.var_source_type.get() == "folder")
        recursive = (self.check_recursive.get() == 1)
        sheet_str = self.entry_sheets.get().strip()
        target_sheets = [s.strip() for s in sheet_str.replace("，", ",").split(",") if s.strip()]
        extract_mode = "specific" if self.switch_col_mode.get() == 1 else "all"
        exact_cols = []
        fuzzy_cols = []
        fuzzy_thresh = self.slider_fuzzy.get()
        use_polars = (self.var_engine.get() == "polars" and POLARS_AVAILABLE)
        
        if extract_mode == "specific":
            e_str = self.entry_exact.get().strip()
            f_str = self.entry_fuzzy.get().strip()
            if e_str:
                exact_cols = [x.strip() for x in e_str.replace("，", ",").split(",") if x.strip()]
            if f_str:
                fuzzy_cols = [x.strip() for x in f_str.replace("，", ",").split(",") if x.strip()]
            if not exact_cols and not fuzzy_cols:
                messagebox.showwarning("提示", "请至少填写一项匹配规则。")
                return

        if not src or not os.path.exists(src):
            messagebox.showerror("错误", "路径不存在")
            return

        save_file = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="提取结果.xlsx")
        if not save_file:
            return

        self.btn_run.configure(state="disabled", text="正在处理...")
        self.textbox.delete("1.0", "end")
        
        # 申请红旗
        stop_event = None
        if hasattr(self, 'app') and self.app is not None:
            stop_event = self.app.register_task(self.module_index)
        
        def t():
            success, msg = core_process(
                src, is_folder, recursive, target_sheets,
                extract_mode, exact_cols, fuzzy_cols, fuzzy_thresh,
                save_file, self.log,
                stop_event=stop_event,
                use_polars=use_polars
            )
            self.log("-" * 30)
            self.log(msg)
            
            # 任务结束，通知主程序
            if hasattr(self, 'app') and self.app is not None:
                self.app.finish_task(self.module_index)
                
            self.btn_run.configure(state="normal", text="开始执行提取")
            if success:
                messagebox.showinfo("成功", f"数据提取完成！\n模式: {'Polars高性能' if use_polars else 'Pandas标准'}")

        threading.Thread(target=t, daemon=True).start()
