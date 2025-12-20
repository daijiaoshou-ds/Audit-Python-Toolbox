import os
import sys
import threading
import csv
import difflib
import re
import urllib.parse
import openpyxl 
from openpyxl import Workbook, load_workbook
import customtkinter as ctk
from tkinter import filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- 1. 引擎加载区 ---
try:
    from python_calamine import CalamineWorkbook
except ImportError:
    CalamineWorkbook = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# COM 接口支持检查
try:
    import win32com.client as win32
    HAS_COM = True
except ImportError:
    HAS_COM = False

# ==============================================================================
#                               核心逻辑区
# ==============================================================================

# --- A. 内容检索逻辑 ---

def clean_val(val):
    if val is None: return ""
    val_str = str(val)
    if val_str.startswith("<") and "Error" in val_str:
        return val_str
    s = val_str.strip()
    if s.endswith(".0"):
        try:
            float(s); s = s[:-2]
        except: pass
    return s

def is_match(content, keyword, threshold):
    c_str = content.lower()
    k_str = keyword.lower()
    if not c_str: return False
    if threshold >= 0.99: return k_str in c_str 
    return difflib.SequenceMatcher(None, k_str, c_str).ratio() >= threshold

# === 【修改点 1】增加 stop_event 参数 ===
def scan_values_rust(file_path, keywords, threshold, stop_event=None):
    """Rust 极速查值"""
    hits = []
    try:
        wb = CalamineWorkbook.from_path(file_path)
        for sheet_name in wb.sheet_names:
            # === 中断检测 ===
            if stop_event and stop_event.is_set(): return hits
            
            try:
                rows = wb.get_sheet_by_name(sheet_name).to_python(skip_empty_area=False)
            except: continue
            for r_idx, row in enumerate(rows):
                for c_idx, cell_value in enumerate(row):
                    val_str = clean_val(cell_value)
                    if not val_str: continue
                    for kw in keywords:
                        if is_match(val_str, kw, threshold):
                            col_letter = openpyxl.utils.get_column_letter(c_idx + 1)
                            hits.append({
                                "file": os.path.basename(file_path),
                                "pos": f"{sheet_name}!{col_letter}{r_idx+1}",
                                "val": val_str
                            })
                            break
    except: pass
    return hits

# === 【修改点 2】增加 stop_event 参数 ===
def scan_values_openpyxl(file_path, keywords, threshold, stop_event=None):
    """OpenPyXL 查值"""
    hits = []
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        for sheet_name in wb.sheetnames:
            # === 中断检测 ===
            if stop_event and stop_event.is_set(): 
                wb.close()
                return hits
                
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    if not cell.value: continue
                    val_str = clean_val(cell.value)
                    if not val_str: continue
                    for kw in keywords:
                        if is_match(val_str, kw, threshold):
                            hits.append({
                                "file": os.path.basename(file_path),
                                "pos": f"{sheet_name}!{cell.coordinate}",
                                "val": val_str
                            })
                            break
        wb.close()
    except: pass
    return hits

# --- B. 外部链接管理逻辑 (保持不变，此处不需要中断) ---

def extract_links_from_file(file_path):
    links_info = []
    try:
        wb = load_workbook(file_path, data_only=False, keep_vba=True, read_only=False)
        base_links = {}
        if hasattr(wb, "_external_links"):
            for idx, link in enumerate(wb._external_links):
                target = "Unknown"
                try: target = link.file_link.target
                except: pass
                base_links[idx + 1] = target 
        found_refs = set()
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    val = cell.value
                    if isinstance(val, str) and val.startswith("=["):
                        match = re.match(r"^=\[(\d+)\](.*?)!", val)
                        if match:
                            idx = int(match.group(1))
                            sheet_part = match.group(2)
                            if sheet_part.startswith("'") and sheet_part.endswith("'"):
                                sheet_part = sheet_part[1:-1]
                            found_refs.add((idx, sheet_part))
        processed_indices = set()
        for idx, sheet_name in found_refs:
            target = base_links.get(idx, "Unknown")
            links_info.append({"index": idx, "target": target, "sheet": sheet_name})
            processed_indices.add(idx)
        for idx, target in base_links.items():
            if idx not in processed_indices:
                links_info.append({"index": idx, "target": target, "sheet": "(未使用)"})
        wb.close()
    except Exception as e:
        return None, str(e)
    return links_info, ""

class ExcelComEngine:
    def __init__(self, log_func):
        self.log = log_func
        self.app = None
        
    def start(self):
        if not HAS_COM:
            self.log("错误: 未检测到 pywin32，无法调用 Excel。请安装: pip install pywin32")
            return False
        try:
            try: self.app = win32.Dispatch("Excel.Application")
            except: self.app = win32.Dispatch("Ket.Application") 
            self.app.Visible = False
            self.app.DisplayAlerts = False 
            return True
        except Exception as e:
            self.log(f"启动 Excel 失败: {e}")
            return False

    def close(self):
        if self.app:
            try: self.app.Quit()
            except: pass
            self.app = None

    def process_file(self, file_path, updates):
        abs_path = os.path.abspath(file_path)
        try:
            wb = self.app.Workbooks.Open(abs_path, UpdateLinks=0)
            current_links = wb.LinkSources(1)
            if not current_links:
                wb.Close(SaveChanges=False)
                return False, "文件内未检测到有效链接源"
            changed = 0
            for idx, new_path in updates.items():
                if idx <= len(current_links):
                    old_link_name = current_links[idx - 1]
                    if new_path and new_path != old_link_name:
                        try:
                            wb.ChangeLink(Name=old_link_name, NewName=new_path, Type=1)
                            changed += 1
                        except Exception as change_err:
                            self.log(f"  - 替换链接失败 (Index {idx}): {change_err}")
            if changed > 0:
                wb.Save()
                wb.Close()
                return True, f"成功更新 {changed} 个链接源"
            else:
                wb.Close(SaveChanges=False)
                return True, "无变化 (可能新路径与原路径相同)"
        except Exception as e:
            return False, f"处理出错: {e}"

# ==============================================================================
#                               界面模块
# ==============================================================================

class keyWordSearchModule:
    def __init__(self):
        self.name = "内容检索&链接管理"
        self.selected_paths = [] 
        self.search_results = []
        self.entry_map_path = None
        self.col_widths_s = [220, 150, 450] 
        # self.app 会由 main.py 注入

    def render(self, parent_frame):
        for widget in parent_frame.winfo_children(): widget.destroy()

        main_scroll = ctk.CTkScrollableFrame(parent_frame, fg_color="transparent", scrollbar_button_color="#E0E0E0", scrollbar_button_hover_color="#D0D0D0")
        main_scroll.pack(fill="both", expand=True, padx=0, pady=0)

        title_frame = ctk.CTkFrame(main_scroll, fg_color="transparent")
        title_frame.pack(fill="x", padx=10, pady=(15, 5))
        ctk.CTkLabel(title_frame, text="🔍 内容检索 & 外部链接管理", font=("Microsoft YaHei", 20, "bold"), text_color="#333").pack(side="left")

        self.tabview = ctk.CTkTabview(main_scroll, fg_color="transparent", segmented_button_fg_color="#F0F0F0", segmented_button_selected_color="#0984e3", segmented_button_selected_hover_color="#0984e3", segmented_button_unselected_color="#E0E0E0", segmented_button_unselected_hover_color="#D6D6D6", text_color="#333", height=600)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        self.tab_search = self.tabview.add("内容/数值检索")
        self.tab_link = self.tabview.add("外部链接替换")

        self.render_search_tab(self.tab_search)
        self.render_link_tab(self.tab_link)

    # ------------------ Tab 1: 内容检索 ------------------
    def render_search_tab(self, parent):
        opt_frame = ctk.CTkFrame(parent, fg_color="white", corner_radius=6, border_width=1, border_color="#E0E0E0")
        opt_frame.pack(fill="x", padx=0, pady=5)

        row1 = ctk.CTkFrame(opt_frame, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=(15, 10))
        ctk.CTkButton(row1, text="+ 文件夹", command=self.add_folder, width=90, fg_color="#F0F5FF", text_color="#007AFF").pack(side="left", padx=5)
        ctk.CTkButton(row1, text="+ 文件", command=self.add_files, width=80, fg_color="#F0F5FF", text_color="#007AFF").pack(side="left", padx=5)
        self.lbl_count = ctk.CTkLabel(row1, text="未选择文件", text_color="#999")
        self.lbl_count.pack(side="left", padx=15)
        ctk.CTkButton(row1, text="清空", command=self.clear_selection, width=60, fg_color="transparent", text_color="#d63031").pack(side="right", padx=5)

        row2 = ctk.CTkFrame(opt_frame, fg_color="transparent")
        row2.pack(fill="x", padx=10, pady=(0, 10))
        self.entry_kw = ctk.CTkEntry(row2, placeholder_text="输入查找内容 (支持多内容，请用 | 分隔，如: 采购|销售|1,000)", height=40)
        self.entry_kw.pack(fill="x")

        row3 = ctk.CTkFrame(opt_frame, fg_color="transparent")
        row3.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkLabel(row3, text="匹配度:", text_color="#333", width=60).pack(side="left")
        self.lbl_slider_val = ctk.CTkLabel(row3, text="1.0", text_color="#0984e3", width=30)
        self.lbl_slider_val.pack(side="left")
        self.slider = ctk.CTkSlider(row3, from_=0.1, to=1.0, number_of_steps=90, width=150, command=lambda v: self.lbl_slider_val.configure(text=f"{v:.1f}"))
        self.slider.set(1.0)
        self.slider.pack(side="left", padx=5)
        
        self.var_rust = ctk.BooleanVar(value=False)
        if CalamineWorkbook: ctk.CTkSwitch(row3, text="Rust极速模式", variable=self.var_rust, text_color="#555").pack(side="right")

        btn_row = ctk.CTkFrame(parent, fg_color="transparent")
        btn_row.pack(fill="x", padx=0, pady=10)
        self.btn_run_search = ctk.CTkButton(btn_row, text="开始检索", command=self.run_search, height=45, fg_color="#0984e3", font=("Microsoft YaHei", 15, "bold"))
        self.btn_run_search.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkButton(btn_row, text="导出结果", command=self.export_search, height=45, fg_color="#00b894", font=("Microsoft YaHei", 15, "bold")).pack(side="right", fill="x", expand=True, padx=(5, 0))

        self.lbl_search_status = ctk.CTkLabel(parent, text="就绪", text_color="#666", anchor="w")
        self.lbl_search_status.pack(fill="x", padx=5, pady=(0, 5))

        res_container = ctk.CTkFrame(parent, fg_color="transparent")
        res_container.pack(fill="both", expand=True, padx=0, pady=5)
        header_grid = ctk.CTkFrame(res_container, fg_color="#E0E0E0", height=30, corner_radius=2)
        header_grid.pack(fill="x")
        headers = ["文件", "位置", "内容"]
        for i, h in enumerate(headers): ctk.CTkLabel(header_grid, text=h, width=self.col_widths_s[i], anchor="w", font=("Arial", 11, "bold")).pack(side="left", padx=10)

        self.res_frame = ctk.CTkFrame(res_container, fg_color="white", corner_radius=0)
        self.res_frame.pack(fill="both", expand=True)
        ctk.CTkLabel(self.res_frame, text="暂无结果", text_color="#CCC", height=50).pack()

    # ------------------ Tab 2: 外部链接管理 ------------------
    def render_link_tab(self, parent):
        step1 = ctk.CTkFrame(parent, fg_color="white", corner_radius=6, border_width=1, border_color="#E0E0E0")
        step1.pack(fill="x", padx=0, pady=10)
        s1_head = ctk.CTkFrame(step1, fg_color="transparent")
        s1_head.pack(fill="x", padx=15, pady=(15, 5))
        ctk.CTkLabel(s1_head, text="1. 扫描导出", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(side="left")
        ctk.CTkLabel(s1_head, text="(需先在【内容检索】页选择文件)", text_color="#999", font=("Microsoft YaHei", 12)).pack(side="left", padx=10)
        ctk.CTkButton(step1, text="导出映射表 (Excel)", command=self.run_link_scan, height=40, fg_color="#0984e3", font=("Microsoft YaHei", 13, "bold")).pack(fill="x", padx=15, pady=(5, 15))

        step2 = ctk.CTkFrame(parent, fg_color="white", corner_radius=6, border_width=1, border_color="#E0E0E0")
        step2.pack(fill="x", padx=0, pady=0)
        s2_head = ctk.CTkFrame(step2, fg_color="transparent")
        s2_head.pack(fill="x", padx=15, pady=(15, 5))
        ctk.CTkLabel(s2_head, text="2. 替换执行", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(side="left")
        ctk.CTkLabel(s2_head, text="(调用 Excel/WPS 进程)", text_color="#d63031", font=("Microsoft YaHei", 12)).pack(side="left", padx=10)
        
        f_in = ctk.CTkFrame(step2, fg_color="transparent")
        f_in.pack(fill="x", padx=15, pady=(0, 10))
        self.entry_map_path = ctk.CTkEntry(f_in, placeholder_text="选择映射表...", height=35)
        self.entry_map_path.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_in, text="浏览", width=70, command=self.select_map_file, height=35, fg_color="#F0F0F0", text_color="#333", hover_color="#DDD").pack(side="right", padx=(5,0))

        self.btn_run_com = ctk.CTkButton(step2, text="启动 Excel 执行替换", command=self.run_link_replace_com, height=40, fg_color="#d63031", hover_color="#b71c1c", font=("Microsoft YaHei", 13, "bold"))
        self.btn_run_com.pack(fill="x", padx=15, pady=(5, 15))

        ctk.CTkLabel(parent, text="执行日志:", text_color="#666", anchor="w", font=("Arial", 12, "bold")).pack(fill="x", padx=5, pady=(15, 5))
        self.log_box = ctk.CTkTextbox(parent, height=180, fg_color="white", border_width=1, border_color="#DDD", text_color="#333", font=("Consolas", 11))
        self.log_box.pack(fill="x", padx=0, pady=(0, 10))

    # ================= 交互逻辑 =================
    
    # ... (辅助函数保持不变) ...
    def add_folder(self):
        d = filedialog.askdirectory()
        if d: self.selected_paths.append(d); self.update_cnt()
    def add_files(self):
        fs = filedialog.askopenfilenames()
        if fs: self.selected_paths.extend(fs); self.update_cnt()
    def clear_selection(self):
        self.selected_paths = []; self.update_cnt()
    def update_cnt(self):
        self.lbl_count.configure(text=f"已选 {len(self.selected_paths)} 项")
    def select_map_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p: self.entry_map_path.delete(0, "end"); self.entry_map_path.insert(0, p)
    def log_link(self, msg):
        self.log_box.insert("end", msg + "\n"); self.log_box.see("end")
    def get_all_files(self):
        files = []
        for p in self.selected_paths:
            if os.path.isfile(p): files.append(p)
            elif os.path.isdir(p):
                for root, _, fs in os.walk(p):
                    for f in fs:
                        if not f.startswith("~$") and f.endswith((".xlsx", ".xlsm")): files.append(os.path.join(root, f))
        return list(set(files))

    # --- Tab 1 任务逻辑 ---
    def run_search(self):
        raw_kw = self.entry_kw.get()
        kws = [k.strip() for k in raw_kw.split('|') if k.strip()]
        if not self.selected_paths or not kws: return messagebox.showwarning("提示", "请选择文件并输入关键词")
        
        # 申请红旗
        stop_event = None
        if hasattr(self, 'app'): stop_event = self.app.register_task(self.module_index)
        self.btn_run_search.configure(state="disabled", text="搜索中...")
        
        for w in self.res_frame.winfo_children(): w.destroy()
        self.lbl_search_status.configure(text="正在初始化...")
        
        files = self.get_all_files()
        threshold = self.slider.get()
        use_rust = self.var_rust.get()
        
        def task():
            results = []
            total_files = len(files)
            aborted = False
            
            for i, f in enumerate(files):
                # === 中断检测 (文件级) ===
                if stop_event and stop_event.is_set():
                    self.lbl_search_status.configure(text=">>> 搜索已被强制终止。")
                    aborted = True
                    break
                    
                self.lbl_search_status.configure(text=f"正在搜索 ({i+1}/{total_files}): {os.path.basename(f)} ...")
                
                # 传入 stop_event 到核心函数
                if use_rust and CalamineWorkbook:
                    res = scan_values_rust(f, kws, threshold, stop_event=stop_event)
                else:
                    res = scan_values_openpyxl(f, kws, threshold, stop_event=stop_event)
                results.extend(res)
            
            if not aborted:
                self.search_results = results
                self.lbl_search_status.configure(text=f"搜索完成。共找到 {len(results)} 条结果。")
                # UI更新部分省略(保持不变)...
                # (为了简洁，这里省略 update_ui 内部代码，因为没有逻辑变更，只在 task 结尾调用)
                self.after_search_complete(results)
            
            # 销假 & 恢复按钮
            if hasattr(self, 'app'): self.app.finish_task(self.module_index)
            self.btn_run_search.configure(state="normal", text="开始检索")
            
        threading.Thread(target=task, daemon=True).start()

    def after_search_complete(self, results):
        # 辅助函数：更新搜索结果 UI
        for w in self.res_frame.winfo_children(): w.destroy()
        if not results: 
            ctk.CTkLabel(self.res_frame, text="未找到匹配项", text_color="#999").pack(pady=20)
            return
        limit = 100 # 限制显示条数避免卡顿
        for i, r in enumerate(results):
            if i >= limit:
                ctk.CTkLabel(self.res_frame, text=f"... 剩余 {len(results)-limit} 条请导出 ...", text_color="#d63031").pack(pady=5)
                break
            row = ctk.CTkFrame(self.res_frame, fg_color="transparent", height=30)
            row.pack(fill="x")
            row.pack_propagate(False)
            f1 = ctk.CTkFrame(row, fg_color="transparent", width=self.col_widths_s[0]); f1.pack(side="left", fill="y"); f1.pack_propagate(False)
            ctk.CTkLabel(f1, text=r['file'], anchor="w", font=("Arial", 11)).pack(side="left", padx=10)
            f2 = ctk.CTkFrame(row, fg_color="transparent", width=self.col_widths_s[1]); f2.pack(side="left", fill="y"); f2.pack_propagate(False)
            ctk.CTkLabel(f2, text=r['pos'], anchor="w", font=("Arial", 11), text_color="#00b894").pack(side="left")
            f3 = ctk.CTkFrame(row, fg_color="transparent", width=self.col_widths_s[2]); f3.pack(side="left", fill="y"); f3.pack_propagate(False)
            ctk.CTkLabel(f3, text=r['val'], anchor="w", font=("Arial", 11)).pack(side="left")
            ctk.CTkFrame(self.res_frame, fg_color="#F0F0F0", height=1).pack(fill="x")

    def export_search(self):
        if not self.search_results: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="检索结果.xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            try:
                wb = Workbook()
                ws = wb.active
                ws.append(["文件", "位置", "内容"])
                for r in self.search_results:
                    val = str(r['val']) if r['val'] else ""
                    val = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', val)
                    ws.append([r['file'], r['pos'], val])
                wb.save(path)
                messagebox.showinfo("完成", "导出成功")
            except Exception as e:
                messagebox.showerror("失败", str(e))

    # --- Tab 2 任务逻辑 ---
    def run_link_scan(self):
        # 扫描很快，不加中断
        files = self.get_all_files()
        if not files: return messagebox.showwarning("提示", "请先在【内容检索】页签添加 Excel 文件")
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="外部链接映射表.xlsx")
        if not save_path: return

        self.log_box.delete("1.0", "end")
        self.log_link("开始扫描外部链接...")

        def task():
            try:
                wb_out = Workbook(); ws_out = wb_out.active
                ws_out.append(["文件全路径", "文件名", "链接索引([N])", "当前链接路径", "原引用Sheet名(参考)", "新链接路径(必填)", "新Sheet名(选填)"])
                count = 0
                for f in files:
                    links, err = extract_links_from_file(f)
                    if err: self.log_link(f"[错] {os.path.basename(f)}: {err}"); continue
                    if links:
                        for l in links: ws_out.append([f, os.path.basename(f), l['index'], l['target'], l['sheet'], "", ""])
                        count += 1
                        self.log_link(f"[扫描] {os.path.basename(f)} 发现 {len(links)} 个引用点")
                wb_out.save(save_path)
                self.log_link(f"-"*30)
                self.log_link(f"扫描结束。共发现含有链接的文件: {count} 个")
                self.entry_map_path.delete(0, "end"); self.entry_map_path.insert(0, save_path)
            except Exception as e:
                self.log_link(f"扫描导出失败: {e}")

        threading.Thread(target=task, daemon=True).start()

    def run_link_replace_com(self):
        map_file = self.entry_map_path.get()
        if not os.path.exists(map_file): return messagebox.showerror("错误", "找不到映射表文件")
        if not messagebox.askyesno("确认", "将启动 Excel 进程执行更改源。\n\n请确保：\n1. 所有 Excel 文件已关闭\n2. 已安装 pywin32\n\n是否继续？"): return

        # 申请红旗
        stop_event = None
        if hasattr(self, 'app'): stop_event = self.app.register_task(self.module_index)
        self.btn_run_com.configure(state="disabled", text="执行中...")

        self.log_link("\n启动 Excel 引擎...")
        
        def task():
            try:
                wb_map = load_workbook(map_file, data_only=True)
                ws_map = wb_map.active
                tasks = {} 
                for row in ws_map.iter_rows(min_row=2, values_only=True):
                    if not row or len(row) < 6: continue
                    f_path, idx, new_path = row[0], row[2], row[5]
                    if new_path and str(new_path).strip():
                        if f_path not in tasks: tasks[f_path] = {}
                        tasks[f_path][int(idx)] = str(new_path).strip()
            except Exception as e:
                self.log_link(f"映射表读取失败: {e}"); return

            if not tasks: self.log_link("未发现任务 (E列为空)"); return

            engine = ExcelComEngine(self.log_link)
            if not engine.start(): return
            
            success_cnt = 0
            for f_path, file_updates in tasks.items():
                # === 中断检测 ===
                if stop_event and stop_event.is_set():
                    self.log_link(">>> 用户强制终止！")
                    break

                if not os.path.exists(f_path): self.log_link(f"[跳过] 文件不存在: {f_path}"); continue
                
                self.log_link(f"正在处理: {os.path.basename(f_path)} ...")
                ok, msg = engine.process_file(f_path, file_updates)
                if ok: self.log_link(f"  -> {msg}"); success_cnt += 1
                else: self.log_link(f"  -> [失败] {msg}")
            
            engine.close()
            self.log_link(f"-"*30)
            self.log_link(f"任务结束。成功: {success_cnt}")
            
            # 销假 & 恢复按钮
            if hasattr(self, 'app'): self.app.finish_task(self.module_index)
            self.btn_run_com.configure(state="normal", text="启动 Excel 执行替换")

        threading.Thread(target=task, daemon=True).start()