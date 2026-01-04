import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
import os
import difflib

# 不在开头导入，放入函数内部
# import torch
# from modules.audit_radar.data_processor import AuditDataProcessor
# from modules.audit_radar.engine import AuditEngine

# --- 风格配置 ---
THEME_COLOR = "#007AFF"
WARN_COLOR = "#d63031"
BORDER_COLOR = "#E5E5E5"
SCROLLBAR_COLOR = "#E0E0E0"
SCROLLBAR_HOVER = "#D0D0D0"

FONT_HEAD = ("Microsoft YaHei", 15, "bold")
FONT_BODY = ("Microsoft YaHei", 12)
FONT_BOLD = ("Microsoft YaHei", 12, "bold")
FONT_LOG = ("Microsoft YaHei", 13) 

class AuditRadarModule:
    def __init__(self):
        self.name = "会计分录测试"
        self.df = None
        self.file_path = ""
        self.chk_vars_amt = {}
        self.chk_vars_cat = {}
        self.filter_keywords = [] 

    def render(self, parent_frame):
        for w in parent_frame.winfo_children(): w.destroy()
        
        self.main_scroll = ctk.CTkScrollableFrame(
            parent_frame, 
            fg_color="transparent",
            scrollbar_button_color=SCROLLBAR_COLOR,
            scrollbar_button_hover_color=SCROLLBAR_HOVER
        )
        self.main_scroll.pack(fill="both", expand=True, padx=10, pady=10)

        header = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(10, 20))

        # 直接写在一起，靠左对齐
        ctk.CTkLabel(header, text="会计分录测试", font=("Microsoft YaHei", 24, "bold"), text_color="#333").pack(side="left")
        
        # 副标题保持在右边一点
        ctk.CTkLabel(header, text="PyTorch 无监督异常检测", font=("Microsoft YaHei", 12), text_color="gray").pack(side="left", padx=10, pady=(10, 0)) 

        self.create_step_frame(self.main_scroll, "Step 1. 导入序时账", self.render_step1)
        self.create_step_frame(self.main_scroll, "Step 2. 训练维度选择", self.render_step2)
        self.create_step_frame(self.main_scroll, "Step 3. 审计参数配置", self.render_step3)
        self.create_step_frame(self.main_scroll, "Step 4. 执行与日志", self.render_step4)

    def create_step_frame(self, parent, title, render_func):
        frame = ctk.CTkFrame(parent, fg_color="white", corner_radius=8, border_width=2, border_color=BORDER_COLOR)
        frame.pack(fill="x", padx=10, pady=8)
        title_lbl = ctk.CTkLabel(frame, text=title, font=FONT_HEAD, text_color=THEME_COLOR)
        title_lbl.pack(anchor="w", padx=15, pady=(12, 5))
        content = ctk.CTkFrame(frame, fg_color="transparent")
        content.pack(fill="x", padx=15, pady=(0, 15))
        render_func(content)

    def render_step1(self, parent):
        self.btn_load = ctk.CTkButton(parent, text="📂 选择 Excel/CSV 文件", command=self.load_file_thread, width=160, height=36, font=FONT_BOLD, fg_color="#F0F5FA", text_color=THEME_COLOR, hover_color="#E1EBF5")
        self.btn_load.pack(side="left")
        self.lbl_file = ctk.CTkLabel(parent, text="未选择文件", text_color="gray", font=FONT_BODY)
        self.lbl_file.pack(side="left", padx=15)
        self.progress = ctk.CTkProgressBar(parent, width=200, mode="indeterminate", height=8)
        self.progress.pack(side="right", padx=10); self.progress.pack_forget()

    def render_step2(self, parent):
        parent.grid_columnconfigure(0, weight=1); parent.grid_columnconfigure(1, weight=1)
        
        f_l = ctk.CTkFrame(parent, fg_color="#F9F9F9", corner_radius=6)
        f_l.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        ctk.CTkLabel(f_l, text="💰 金额维度 (借/贷/金额)", font=FONT_BOLD, text_color="#333").pack(anchor="w", padx=10, pady=5)
        self.scroll_amt = ctk.CTkScrollableFrame(f_l, height=120, fg_color="white", scrollbar_button_color=SCROLLBAR_COLOR, scrollbar_button_hover_color=SCROLLBAR_HOVER)
        self.scroll_amt.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkLabel(self.scroll_amt, text="请先加载文件...", text_color="gray").pack(pady=10)

        f_r = ctk.CTkFrame(parent, fg_color="#F9F9F9", corner_radius=6)
        f_r.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        ctk.CTkLabel(f_r, text="🏷️ 特征维度 (科目/明细/摘要)", font=FONT_BOLD, text_color="#333").pack(anchor="w", padx=10, pady=5)
        self.scroll_cat = ctk.CTkScrollableFrame(f_r, height=120, fg_color="white", scrollbar_button_color=SCROLLBAR_COLOR, scrollbar_button_hover_color=SCROLLBAR_HOVER)
        self.scroll_cat.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkLabel(self.scroll_cat, text="请先加载文件...", text_color="gray").pack(pady=10)

    def render_step3(self, parent):
        # --- 第一行：基础参数 ---
        r1 = ctk.CTkFrame(parent, fg_color="transparent"); r1.pack(fill="x", pady=5)
        
        # 重要性水平
        ctk.CTkLabel(r1, text="重要性水平:", text_color="#333", width=80, anchor="w", font=FONT_BODY).pack(side="left")
        self.entry_threshold = ctk.CTkEntry(r1, width=90, border_color="#CCC"); self.entry_threshold.insert(0, "0"); self.entry_threshold.pack(side="left", padx=5)
        ctk.CTkLabel(r1, text="元", text_color="gray").pack(side="left", padx=(0, 5))
        
        # === 【新增】测算按钮 ===
        ctk.CTkButton(r1, text="🔍 测算过滤量", command=self.calculate_threshold_stats, width=90, height=28, fg_color="#F0F5FA", text_color=THEME_COLOR, hover_color="#E1EBF5").pack(side="left", padx=5)

        # 训练轮数
        ctk.CTkLabel(r1, text="训练轮数:", text_color="#333", width=70, anchor="w", font=FONT_BODY).pack(side="left", padx=(20, 0))
        self.slider_epoch = ctk.CTkSlider(r1, from_=50, to=300, number_of_steps=5, width=120, command=self.update_epoch_label); self.slider_epoch.set(150); self.slider_epoch.pack(side="left", padx=5)
        self.lbl_epoch = ctk.CTkLabel(r1, text="150 轮", text_color=THEME_COLOR, font=FONT_BOLD, width=50); self.lbl_epoch.pack(side="left", padx=5)

        ctk.CTkFrame(parent, height=2, fg_color="#F0F0F0").pack(fill="x", pady=10)
        ctk.CTkLabel(parent, text="🧹 摘要关键词过滤 (排除无意义分录，如结转损益)", font=FONT_BOLD, text_color="#333").pack(anchor="w", pady=(0, 5))
        r2 = ctk.CTkFrame(parent, fg_color="transparent"); r2.pack(fill="x")
        self.entry_filter = ctk.CTkEntry(r2, placeholder_text="输入关键词，如：结转期间损益", width=200); self.entry_filter.pack(side="left")
        ctk.CTkButton(r2, text="+ 添加", width=60, command=self.add_filter_keyword, fg_color="#F0F5FA", text_color="#333", hover_color="#DDD").pack(side="left", padx=5)
        ctk.CTkButton(r2, text="清空", width=60, command=self.clear_filter_keywords, fg_color="transparent", text_color="red", hover_color="#FEE").pack(side="left")
        ctk.CTkLabel(r2, text="匹配严格度:", text_color="#555").pack(side="left", padx=(20, 5))
        self.slider_sim = ctk.CTkSlider(r2, from_=0.1, to=1.0, number_of_steps=9, width=100, command=self.update_sim_label); self.slider_sim.set(0.6); self.slider_sim.pack(side="left")
        self.lbl_sim = ctk.CTkLabel(r2, text="0.6", text_color=THEME_COLOR, font=("Arial", 12, "bold"), width=30); self.lbl_sim.pack(side="left", padx=5)
        self.scroll_filter = ctk.CTkScrollableFrame(parent, height=60, fg_color="#F9F9F9", orientation="horizontal", scrollbar_button_color=SCROLLBAR_COLOR, scrollbar_button_hover_color=SCROLLBAR_HOVER); self.scroll_filter.pack(fill="x", pady=(5, 0))
        self.lbl_no_filter = ctk.CTkLabel(self.scroll_filter, text="暂无过滤词", text_color="gray", font=("Arial", 11)); self.lbl_no_filter.pack(pady=5)
        self.add_filter_keyword("期间损益结转") 

    def render_step4(self, parent):
        self.btn_run = ctk.CTkButton(parent, text="🚀 启动雷达扫描", command=self.run_analysis, height=45, font=("Microsoft YaHei", 16, "bold"), fg_color=WARN_COLOR)
        self.btn_run.pack(fill="x", pady=(5, 15))
        header = ctk.CTkFrame(parent, fg_color="#F0F0F0", height=28, corner_radius=4)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="  运行日志 (Console)", font=("Arial", 11, "bold"), text_color="#666").place(rely=0.5, anchor="w")
        self.log_box = ctk.CTkTextbox(parent, height=300, fg_color="#FAFAFA", text_color="#333", font=FONT_LOG, border_color="#DDD", border_width=1, corner_radius=0, scrollbar_button_color=SCROLLBAR_COLOR, scrollbar_button_hover_color=SCROLLBAR_HOVER)
        self.log_box.pack(fill="x")

    # ... (辅助函数保持不变) ...
    def update_epoch_label(self, val): self.lbl_epoch.configure(text=f"{int(val)} 轮")
    def update_sim_label(self, val): self.lbl_sim.configure(text=f"{val:.1f}")
    def log(self, msg): self.log_box.insert("end", msg + "\n"); self.log_box.see("end")
    def add_filter_keyword(self, val=None):
        kw = val if val else self.entry_filter.get().strip()
        if not kw: return
        if kw not in self.filter_keywords: self.filter_keywords.append(kw); self.refresh_filter_ui()
        self.entry_filter.delete(0, "end")
    def clear_filter_keywords(self): self.filter_keywords = []; self.refresh_filter_ui()
    def delete_filter_keyword(self, kw):
        if kw in self.filter_keywords: self.filter_keywords.remove(kw); self.refresh_filter_ui()
    def refresh_filter_ui(self):
        for w in self.scroll_filter.winfo_children(): w.destroy()
        if not self.filter_keywords: ctk.CTkLabel(self.scroll_filter, text="暂无过滤词", text_color="gray").pack(pady=5); return
        for kw in self.filter_keywords:
            f = ctk.CTkFrame(self.scroll_filter, fg_color="#E1EBF5", corner_radius=10); f.pack(side="left", padx=5, pady=2)
            ctk.CTkLabel(f, text=kw, text_color=THEME_COLOR).pack(side="left", padx=(10, 5))
            ctk.CTkButton(f, text="×", width=20, height=20, fg_color="transparent", text_color="red", hover_color="#D1DBE5", command=lambda k=kw: self.delete_filter_keyword(k)).pack(side="left", padx=(0, 5))
    def load_file_thread(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if not p: return
        self.btn_load.configure(state="disabled", text="读取中...")
        self.progress.pack(side="right", padx=10); self.progress.start()
        def _read():
            try:
                if p.endswith('.csv'): df = pd.read_csv(p)
                else: df = pd.read_excel(p)
                if hasattr(self, 'lbl_file'): self.lbl_file.after(0, lambda: self.on_file_loaded(p, df))
            except Exception as e:
                if hasattr(self, 'lbl_file'): self.lbl_file.after(0, lambda: self.on_load_error(str(e)))
        threading.Thread(target=_read, daemon=True).start()
    def on_file_loaded(self, path, df):
        self.df = df; self.file_path = path
        self.lbl_file.configure(text=os.path.basename(path), text_color="#333")
        self.btn_load.configure(state="normal", text="📂 重新选择")
        self.progress.stop(); self.progress.pack_forget()
        self.log(f"文件加载成功: {os.path.basename(path)} | 行数: {len(df)}")
        self.refresh_columns()
    def on_load_error(self, err_msg):
        self.btn_load.configure(state="normal", text="📂 选择文件")
        self.progress.stop(); self.progress.pack_forget()
        self.lbl_file.configure(text="读取失败", text_color="red"); messagebox.showerror("读取错误", err_msg)
    def refresh_columns(self):
        for w in self.scroll_amt.winfo_children(): w.destroy()
        for w in self.scroll_cat.winfo_children(): w.destroy()
        self.chk_vars_amt = {}; self.chk_vars_cat = {}
        cols = list(self.df.columns)
        def is_default_amt(name): return any(k in str(name) for k in ['借方', '贷方', 'debit', 'credit'])
        def is_default_cat(name): return '科目' in str(name)
        for col in cols:
            var = ctk.BooleanVar(value=is_default_amt(col))
            ctk.CTkCheckBox(self.scroll_amt, text=str(col), variable=var, text_color="#333", font=FONT_BODY).pack(anchor="w", pady=2, padx=5)
            self.chk_vars_amt[col] = var
        for col in cols:
            var = ctk.BooleanVar(value=is_default_cat(col))
            ctk.CTkCheckBox(self.scroll_cat, text=str(col), variable=var, text_color="#333", font=FONT_BODY).pack(anchor="w", pady=2, padx=5)
            self.chk_vars_cat[col] = var
    def get_selected_cols(self):
        amt = [c for c, v in self.chk_vars_amt.items() if v.get()]
        cat = [c for c, v in self.chk_vars_cat.items() if v.get()]
        return amt, cat

    # === 【新增】测算按钮逻辑 ===
    def calculate_threshold_stats(self):
        if self.df is None: return messagebox.showwarning("提示", "请先加载文件")
        amt_cols, _ = self.get_selected_cols()
        if not amt_cols: return messagebox.showwarning("提示", "请先在 Step 2 勾选金额列")
        
        try:
            threshold = float(self.entry_threshold.get())
        except: return messagebox.showwarning("提示", "请输入有效的数字阈值")

        # 计算逻辑
        self.log(f"--- 测算: 阈值 {threshold} ---")
        
        # 填充0，取绝对值，计算每行的最大值
        max_abs = self.df[amt_cols].abs().max(axis=1).fillna(0)
        
        # 找出小于阈值的行
        mask_ignored = max_abs < threshold
        count_ignored = mask_ignored.sum()
        total_count = len(self.df)
        ratio = (count_ignored / total_count) * 100
        
        ignored_df = self.df[mask_ignored]
        
        self.log(f"📉 预计过滤统计:")
        self.log(f"   总行数: {total_count}")
        self.log(f"   将被忽略: {count_ignored} 行 ({ratio:.1f}%)")
        self.log(f"   剩余有效: {total_count - count_ignored} 行")
        
        # 尝试统计金额
        for col in amt_cols:
            col_sum = ignored_df[col].sum()
            self.log(f"   [{col}] 被忽略总额: {col_sum:,.2f}")
        
        self.log("-" * 20)

    # === 执行逻辑 ===
    def run_analysis(self):
        if self.df is None: return messagebox.showwarning("提示", "请加载文件")
        amt_cols, cat_cols = self.get_selected_cols()
        if not amt_cols or not cat_cols: return messagebox.showwarning("提示", "请至少选择一列金额和一列特征")
        try:
            threshold = float(self.entry_threshold.get())
            epochs = int(self.slider_epoch.get())
            sim_threshold = self.slider_sim.get()
        except: return messagebox.showwarning("提示", "参数格式错误")

        stop_event = None
        if hasattr(self, 'app'): stop_event = self.app.register_task(self.module_index)
        self.btn_run.configure(state="disabled", text="AI 计算中...")
        self.log_box.delete("1.0", "end")
        
        def task():
            try:
                # ====== 导入pytorch ======
                import torch
                from modules.audit_radar.data_processor import AuditDataProcessor
                from modules.audit_radar.engine import AuditEngine
                # ==============
                self.log(f"特征列: {cat_cols}")
                self.log(f"阈值: {threshold} | 轮数: {epochs}")
                if self.filter_keywords: self.log(f"启用摘要过滤: {self.filter_keywords} (严格度: {sim_threshold:.1f})")
                
                processor = AuditDataProcessor()
                self.log("数据清洗与预处理...")
                processed_df = processor.preprocess(self.df, amt_cols, cat_cols)
                
                device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
                cat_t, cont_t = processor.get_tensors(processed_df, device)
                
                engine = AuditEngine(processor, device)
                self.log("开始训练自编码器...")
                success, final_loss = engine.train_model(cat_t, cont_t, epochs=epochs, log_callback=self.log, stop_event=stop_event)
                if not success: return

                # 诊断
                self.log("-" * 30)
                self.log(f"📢 模型诊断报告 (Final Loss: {final_loss:.4f})")
                if final_loss > 1.0: self.log("🔴 状态：欠拟合 (模型没学会)\n💡 建议：增加训练轮数 (>200)")
                elif final_loss < 0.1: self.log("🔵 状态：过拟合 (模型死记硬背)\n💡 建议：减少训练轮数")
                else: self.log("🟢 状态：黄金区间 (最佳状态)\n💡 说明：模型已掌握核心规律，且保持了对异常的敏感度")
                self.log("-" * 30)

                self.log("正在评分与归因分析...")
                scores, reasons = engine.predict_with_reason(cat_t, cont_t, raw_df=self.df, amt_cols=amt_cols, threshold=threshold)
                
                # 摘要过滤
                if self.filter_keywords and '摘要' in self.df.columns:
                    self.log("正在执行摘要过滤...")
                    filtered_count = 0
                    abstracts = self.df['摘要'].astype(str).str.lower().fillna("")
                    for idx, txt in enumerate(abstracts):
                        if scores[idx] == 0: continue
                        is_match = False
                        for kw in self.filter_keywords:
                            kw_clean = kw.lower()
                            if sim_threshold >= 0.9:
                                if kw_clean in txt: is_match = True; break
                            else:
                                if kw_clean in txt: is_match = True; break
                                ratio = difflib.SequenceMatcher(None, kw_clean, txt).ratio()
                                if ratio >= sim_threshold: is_match = True; break
                        if is_match:
                            scores[idx] = 0.0; reasons[idx] = "忽略(摘要过滤)"; filtered_count += 1
                    self.log(f"  -> 已根据摘要过滤掉 {filtered_count} 条记录")

                self.df["异常评分"] = scores
                if scores.max() > scores.min(): self.df["异常评分"] = (scores - scores.min()) / (scores.max() - scores.min()) * 100
                else: self.df["异常评分"] = 0
                
                self.df["异常评分"] = self.df["异常评分"].apply(lambda x: round(x, 2))
                self.df["异常主要原因"] = reasons
                self.df = self.df.sort_values("异常评分", ascending=False)
                
                # === 【修改点】文件名优化 (去除原扩展名，覆盖旧文件) ===
                # 原文件名: data.xlsx -> 新文件名: data_审计雷达报告.xlsx
                base_name = os.path.splitext(self.file_path)[0]
                out_path = f"{base_name}_审计雷达报告.xlsx"
                
                self.df.head(100000).to_excel(out_path, index=False)
                
                self.log("-" * 30)
                self.log(f"分析完成！结果已保存: {os.path.basename(out_path)}")
                messagebox.showinfo("完成", "扫描结束")
                os.startfile(os.path.dirname(out_path))

            except Exception as e:
                self.log(f"Error: {e}"); import traceback; print(traceback.format_exc())
            finally:
                if hasattr(self, 'app'): self.app.finish_task(self.module_index)
                self.btn_run.configure(state="normal", text="🚀 启动雷达扫描")

        threading.Thread(target=task, daemon=True).start()