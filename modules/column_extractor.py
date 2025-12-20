import os
import pandas as pd
from openpyxl import load_workbook
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from difflib import SequenceMatcher

# ==================== 核心逻辑层 ====================

def calculate_similarity(s1, s2):
    """计算两个字符串的相似度"""
    if not s1 or not s2: return 0.0
    return SequenceMatcher(None, s1.lower(), s2.lower()).ratio()

def is_fuzzy_match(target_keyword, cell_value, threshold):
    """智能模糊匹配判断"""
    t = str(target_keyword).strip().lower()
    v = str(cell_value).strip().lower()
    if threshold >= 0.99: return t == v
    if threshold <= 0.9:
        if t in v or v in t: return True
    score = calculate_similarity(t, v)
    return score >= threshold

def get_files_to_process(source_path, is_folder, recursive):
    """根据设置获取文件列表"""
    files_list = []
    if not is_folder:
        if os.path.isfile(source_path) and source_path.lower().endswith(('.xlsx', '.xlsm')):
            return [source_path]
        return []
    if not os.path.exists(source_path): return []

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

def scan_header_and_map_columns(worksheet, mode, exact_cols=None, fuzzy_cols=None, fuzzy_threshold=0.6):
    """扫描表头并建立映射"""
    max_scan = min(worksheet.max_row, 15)
    if max_scan == 0: return {}
    best_mapping = {}
    max_matches = -1
    
    for r in range(1, max_scan + 1):
        row_values = []
        for c in range(1, worksheet.max_column + 1):
            val = worksheet.cell(row=r, column=c).value
            row_values.append(str(val).strip() if val is not None else "")

        if all(v == "" for v in row_values): continue
        current_mapping = {}
        
        if mode == 'all':
            name_counter = {}
            for c_idx, val in enumerate(row_values, 1):
                if not val: continue
                count = name_counter.get(val, 0)
                final_name = val if count == 0 else f"{val}_{count}"
                name_counter[val] = count + 1
                current_mapping[c_idx] = final_name
            if current_mapping: return current_mapping

        else:
            matches_count = 0
            name_counter = {}
            for c_idx, val in enumerate(row_values, 1):
                if not val: continue
                val_lower = val.lower()
                matched_name = None
                for target in exact_cols:
                    if target.lower() == val_lower:
                        matched_name = target
                        break
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

# === 【修改点 1】新增 stop_event 参数 ===
def core_process(source_path, is_folder, recursive, target_sheets, 
                 extract_mode, exact_cols, fuzzy_cols, fuzzy_threshold,
                 save_path, log_func, stop_event=None): 
    
    log_func("=== 任务开始 ===")
    
    files = get_files_to_process(source_path, is_folder, recursive)
    if not files: return False, "未找到有效的 Excel 文件。"
    
    log_func(f"共发现 {len(files)} 个文件，准备处理...")
    
    all_rows = []
    processed_count = 0

    for i, file_path in enumerate(files):
        # === 【修改点 2】循环开始时检查中断 ===
        if stop_event and stop_event.is_set():
            log_func(">>> 用户强制停止任务！")
            return False, "任务已终止。"
        # ===================================

        fname = os.path.basename(file_path)
        log_func(f"[{i+1}/{len(files)}] 正在读取: {fname}")
        
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            
            for ws in wb.worksheets:
                # === 【修改点 3】Sheet循环里也检查，提高响应速度 ===
                if stop_event and stop_event.is_set():
                    wb.close()
                    log_func(">>> 用户强制停止任务！")
                    return False, "任务已终止。"
                # ===============================================

                if target_sheets:
                    is_match = False
                    for target in target_sheets:
                        if is_fuzzy_match(target, ws.title, fuzzy_threshold):
                            is_match = True
                            break
                    if not is_match: continue
                
                col_map = scan_header_and_map_columns(ws, extract_mode, exact_cols, fuzzy_cols, fuzzy_threshold)
                if not col_map: continue

                for row_data in ws.iter_rows(values_only=True):
                    extracted_row = {}
                    has_valid_data = False
                    for col_idx, target_name in col_map.items():
                        if col_idx - 1 < len(row_data):
                            val = row_data[col_idx - 1]
                            extracted_row[target_name] = val
                            if val is not None: has_valid_data = True
                    
                    if has_valid_data:
                        match_header_count = sum(1 for k, v in extracted_row.items() if str(v) == k)
                        if match_header_count > 0 and match_header_count > len(extracted_row) / 2: continue
                        extracted_row['_来源工作簿'] = fname
                        extracted_row['_来源工作表'] = ws.title
                        all_rows.append(extracted_row)
            wb.close()
            processed_count += 1
            
        except Exception as e:
            log_func(f"读取失败 {fname}: {e}")

    if not all_rows: return False, "未提取到任何数据。"

    try:
        log_func("正在汇总并保存...")
        df = pd.DataFrame(all_rows)
        cols = [c for c in df.columns if c not in ['_来源工作簿', '_来源工作表']]
        if extract_mode == 'specific':
            sorted_cols = []
            for exact in exact_cols:
                if exact in cols: sorted_cols.append(exact)
            remaining = [c for c in cols if c not in sorted_cols]
            final_cols = sorted_cols + remaining + ['_来源工作簿', '_来源工作表']
        else:
            final_cols = cols + ['_来源工作簿', '_来源工作表']
            
        df = df.reindex(columns=final_cols)
        df.to_excel(save_path, index=False)
        return True, f"完成！共提取 {len(df)} 行数据。\n保存至: {save_path}"
    except Exception as e:
        return False, f"保存失败: {e}"


# ==================== 界面层 ====================

class ColumnExtractorModule:
    def __init__(self):
        self.name = "Excel数据提取"
        self.input_path = ""
        # 注意：self.app 和 self.module_index 会由 main.py 自动注入
        
    def render(self, parent_frame):
        for widget in parent_frame.winfo_children(): widget.destroy()

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

        self.btn_run = ctk.CTkButton(scroll, text="开始执行提取", height=50, fg_color="#007AFF", font=("Microsoft YaHei", 16, "bold"), command=self.run_task)
        self.btn_run.pack(fill="x", padx=20, pady=20)

        ctk.CTkLabel(scroll, text="执行日志", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(anchor="w", padx=20, pady=(0, 5))
        self.textbox = ctk.CTkTextbox(scroll, height=180, fg_color="#FAFAFA", border_width=1, border_color="#D0D0D0", text_color="#333")
        self.textbox.pack(fill="x", padx=20, pady=(0, 20))

        self.toggle_source_ui()
        self.toggle_col_ui()

    def update_slider_label(self, value):
        val = round(value, 2)
        desc = ""
        if val >= 0.99: desc = "(精确)"
        elif val >= 0.8: desc = "(严格)"
        elif val <= 0.4: desc = "(宽松)"
        else: desc = "(标准)"
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
        
        if extract_mode == "specific":
            e_str = self.entry_exact.get().strip()
            f_str = self.entry_fuzzy.get().strip()
            if e_str: exact_cols = [x.strip() for x in e_str.replace("，", ",").split(",") if x.strip()]
            if f_str: fuzzy_cols = [x.strip() for x in f_str.replace("，", ",").split(",") if x.strip()]
            if not exact_cols and not fuzzy_cols:
                messagebox.showwarning("提示", "请至少填写一项匹配规则。")
                return

        if not src or not os.path.exists(src):
            messagebox.showerror("错误", "路径不存在")
            return

        save_file = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="提取结果.xlsx")
        if not save_file: return

        self.btn_run.configure(state="disabled", text="正在处理...")
        self.textbox.delete("1.0", "end")
        
        # === 【修改点 4】向主程序申请红旗，并传给子线程 ===
        stop_event = None
        if hasattr(self, 'app'): # 确保 app 已注入
            stop_event = self.app.register_task(self.module_index)
        
        def t():
            success, msg = core_process(
                src, is_folder, recursive, target_sheets,
                extract_mode, exact_cols, fuzzy_cols, fuzzy_thresh,
                save_file, self.log,
                stop_event=stop_event  # 传入红旗
            )
            self.log("-" * 30)
            self.log(msg)
            
            # 任务结束，通知主程序销假
            if hasattr(self, 'app'):
                self.app.finish_task(self.module_index)
                
            self.btn_run.configure(state="normal", text="开始执行提取")
            if success: messagebox.showinfo("成功", "数据提取完成！")

        threading.Thread(target=t, daemon=True).start()