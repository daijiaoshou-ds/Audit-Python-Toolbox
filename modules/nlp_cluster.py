import os
import re
import threading
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from modules.path_manager import get_model_path
from difflib import SequenceMatcher

# 取消开头加载，放入函数内部
# from sklearn.cluster import KMeans
# from sklearn.metrics import silhouette_score
# import jieba.analyse


# --- 注意：这里不再导入 sentence_transformers，防止启动卡顿 ---
# try:
#     from sentence_transformers import SentenceTransformer
# except ImportError:
#     pass

# --- 样式常量 ---
FONT_TITLE = ("Microsoft YaHei", 20, "bold")
FONT_SUBTITLE = ("Microsoft YaHei", 14, "bold")
FONT_BODY = ("Microsoft YaHei", 12)
THEME_COLOR = "#007AFF"
BG_CARD = "white"
SCROLLBAR_COLOR = "#E0E0E0"
SCROLLBAR_HOVER = "#D0D0D0"

# --- 默认审计停用词 ---
DEFAULT_STOPWORDS = [
    "凭证", "摘要", "备注", "附件", "入账", "分录", "记账", "核算", "业务", "事项"
]

class NLPClusterModule:
    def __init__(self):
        self.name = "摘要语义聚类分析"
        self.df = None          
        self.df_clean = None    
        self.file_path = ""
        self.model_name = "text2vec-base-chinese"
        
        self.filter_keywords = [] 
        self.label_stopwords = set(DEFAULT_STOPWORDS)
        self.filter_strictness = 0.6
        self.label_topk = 2 

    def render(self, parent_frame):
        for widget in parent_frame.winfo_children(): widget.destroy()

        scroll = ctk.CTkScrollableFrame(
            parent_frame, 
            fg_color="transparent",
            scrollbar_button_color=SCROLLBAR_COLOR,
            scrollbar_button_hover_color=SCROLLBAR_HOVER
        )
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        header = ctk.CTkFrame(scroll, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(10, 15))
        ctk.CTkLabel(header, text="AI审计摘要分析", font=FONT_TITLE, text_color="#333").pack(side="left")
        
        self.create_section_frame(scroll, "Step 1. 数据导入", self.render_import_section)
        self.create_section_frame(scroll, "Step 2. 审计过滤 (数据降噪)", self.render_filter_section)
        self.create_section_frame(scroll, "Step 3. 聚类与标签策略", self.render_cluster_section)
        self.render_run_section(scroll)

    def create_section_frame(self, parent, title, render_func):
        frame = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=8, border_color="#E0E0E0", border_width=1)
        frame.pack(fill="x", padx=10, pady=(0, 10))
        title_bar = ctk.CTkFrame(frame, fg_color="transparent", height=28)
        title_bar.pack(fill="x", padx=15, pady=(10, 5))
        ctk.CTkLabel(title_bar, text=title, font=FONT_SUBTITLE, text_color=THEME_COLOR).pack(side="left")
        ctk.CTkFrame(frame, height=2, fg_color="#F5F5F5").pack(fill="x", padx=15, pady=(0, 5))
        content = ctk.CTkFrame(frame, fg_color="transparent")
        content.pack(fill="x", padx=15, pady=(0, 10))
        render_func(content)

    def render_import_section(self, parent):
        r1 = ctk.CTkFrame(parent, fg_color="transparent")
        r1.pack(fill="x")
        self.btn_file = ctk.CTkButton(r1, text="📂 选择序时账 (Excel)", command=self.load_file_thread, width=160, height=32, font=FONT_BODY)
        self.btn_file.pack(side="left")
        self.lbl_file_status = ctk.CTkLabel(r1, text="未选择文件", text_color="#999", font=FONT_BODY)
        self.lbl_file_status.pack(side="left", padx=15)

        r2 = ctk.CTkFrame(parent, fg_color="transparent")
        r2.pack(fill="x", pady=(10, 0))
        r2.grid_columnconfigure(0, weight=1); r2.grid_columnconfigure(1, weight=1)
        r2.grid_columnconfigure(2, weight=1); r2.grid_columnconfigure(3, weight=1)

        combo_style = {
            "fg_color": "white", "button_color": "#F0F0F0", "button_hover_color": "#E0E0E0",
            "border_color": "#D0D0D0", "text_color": "#333", 
            "dropdown_fg_color": "white", "dropdown_text_color": "#333", "dropdown_hover_color": "#F0F5FA",
            "height": 28
        }

        ctk.CTkLabel(r2, text="摘要列 *", text_color="gray", font=("Arial", 11)).grid(row=0, column=0, sticky="w", padx=5)
        self.combo_abstract = ctk.CTkComboBox(r2, values=[], state="disabled", **combo_style)
        self.combo_abstract.grid(row=1, column=0, sticky="ew", padx=5)

        ctk.CTkLabel(r2, text="一级科目列 (分组) *", text_color="gray", font=("Arial", 11)).grid(row=0, column=1, sticky="w", padx=5)
        self.combo_subject = ctk.CTkComboBox(r2, values=[], state="disabled", **combo_style)
        self.combo_subject.grid(row=1, column=1, sticky="ew", padx=5)
        
        ctk.CTkLabel(r2, text="借方金额", text_color="gray", font=("Arial", 11)).grid(row=0, column=2, sticky="w", padx=5)
        self.combo_debit = ctk.CTkComboBox(r2, values=[], state="disabled", **combo_style)
        self.combo_debit.grid(row=1, column=2, sticky="ew", padx=5)
        
        ctk.CTkLabel(r2, text="贷方金额", text_color="gray", font=("Arial", 11)).grid(row=0, column=3, sticky="w", padx=5)
        self.combo_credit = ctk.CTkComboBox(r2, values=[], state="disabled", **combo_style)
        self.combo_credit.grid(row=1, column=3, sticky="ew", padx=5)

    def render_filter_section(self, parent):
        r0 = ctk.CTkFrame(parent, fg_color="transparent"); r0.pack(fill="x", pady=(0, 5))
        self.var_remove_num = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(r0, text="强力降噪 (剔除所有数字、日期，仅保留中文语义)", variable=self.var_remove_num, text_color="#333", font=("Microsoft YaHei", 12, "bold"), fg_color=THEME_COLOR, height=20).pack(anchor="w")

        r1 = ctk.CTkFrame(parent, fg_color="transparent"); r1.pack(fill="x", pady=(5, 5))
        ctk.CTkLabel(r1, text="重要性水平:", text_color="#333", font=FONT_BODY).pack(side="left")
        self.entry_threshold = ctk.CTkEntry(r1, width=90, height=28, border_color="#CCC"); self.entry_threshold.insert(0, "0"); self.entry_threshold.pack(side="left", padx=10)
        ctk.CTkLabel(r1, text="元", text_color="gray").pack(side="left")
        
        ctk.CTkButton(r1, text="🔍 测算过滤量", command=self.calculate_stats, width=110, height=28, fg_color="#F0F5FA", text_color=THEME_COLOR, hover_color="#E1EBF5").pack(side="right")

        ctk.CTkLabel(parent, text="行过滤关键词 (包含这些词的行将被整行剔除):", text_color="#333", font=FONT_BODY).pack(anchor="w", pady=(5, 2))
        r2 = ctk.CTkFrame(parent, fg_color="transparent"); r2.pack(fill="x")
        self.entry_kw = ctk.CTkEntry(r2, placeholder_text="如: 结转", width=150, height=28); self.entry_kw.pack(side="left")
        ctk.CTkButton(r2, text="+ 添加", width=60, height=28, command=self.add_keyword, fg_color="#EEE", text_color="#333", hover_color="#DDD").pack(side="left", padx=5)
        ctk.CTkButton(r2, text="清空", width=60, height=28, command=self.clear_keywords, fg_color="transparent", text_color="red", hover_color="#FEE").pack(side="left")

        self.frame_keywords = ctk.CTkScrollableFrame(
            parent, height=40, orientation="horizontal", fg_color="#F9F9F9",
            scrollbar_button_color=SCROLLBAR_COLOR, scrollbar_button_hover_color=SCROLLBAR_HOVER
        )
        self.frame_keywords.pack(fill="x", pady=(5, 0))
        self.refresh_keywords_ui()

        self.lbl_stats = ctk.CTkLabel(parent, text="", text_color="#666", font=("Consolas", 12), justify="left", anchor="w")
        self.lbl_stats.pack(pady=(5, 0), anchor="w", fill="x")

    def render_cluster_section(self, parent):
        ctk.CTkLabel(parent, text="A. 聚类参数", text_color="#007AFF", font=("Microsoft YaHei", 12, "bold")).pack(anchor="w", pady=(0, 2))
        
        self.var_split_subject = ctk.BooleanVar(value=True)
        cb_split = ctk.CTkCheckBox(parent, text="启用科目分割聚类 (推荐)", variable=self.var_split_subject, text_color="#333", font=FONT_BODY, height=20)
        cb_split.pack(anchor="w", pady=(0, 5))
        
        r1 = ctk.CTkFrame(parent, fg_color="transparent"); r1.pack(fill="x")
        ctk.CTkLabel(r1, text="聚类粒度:", text_color="#333", font=FONT_BODY).pack(side="left")
        self.lbl_granularity = ctk.CTkLabel(r1, text="适中", text_color=THEME_COLOR, font=("Arial", 12, "bold")); self.lbl_granularity.pack(side="right", padx=10)
        
        self.slider_g = ctk.CTkSlider(r1, from_=1, to=100, number_of_steps=99, command=self.update_granularity_label)
        self.slider_g.set(50); self.slider_g.pack(fill="x", padx=10)
        
        self.lbl_k_recommend = ctk.CTkLabel(parent, text="", text_color="#00b894", font=("Arial", 11))
        self.lbl_k_recommend.pack(anchor="w", pady=(2, 5))

        ctk.CTkFrame(parent, height=1, fg_color="#EEE").pack(fill="x", pady=5)

        ctk.CTkLabel(parent, text="B. 标签生成优化", text_color="#007AFF", font=("Microsoft YaHei", 12, "bold")).pack(anchor="w", pady=(5, 2))
        
        r_topk = ctk.CTkFrame(parent, fg_color="transparent"); r_topk.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(r_topk, text="标签关键词数量:", text_color="#333", font=FONT_BODY).pack(side="left")
        self.lbl_topk = ctk.CTkLabel(r_topk, text="2 个", text_color=THEME_COLOR, font=("Arial", 12, "bold")); self.lbl_topk.pack(side="right", padx=10)
        self.slider_topk = ctk.CTkSlider(r_topk, from_=1, to=5, number_of_steps=4, width=150, command=lambda v: self.lbl_topk.configure(text=f"{int(v)} 个"))
        self.slider_topk.set(2); self.slider_topk.pack(side="right", padx=10)

        r_stop = ctk.CTkFrame(parent, fg_color="transparent"); r_stop.pack(fill="x")
        self.entry_stop = ctk.CTkEntry(r_stop, placeholder_text="如: 招行", width=120, height=28); self.entry_stop.pack(side="left")
        ctk.CTkButton(r_stop, text="+ 停用", width=60, height=28, command=self.add_stopword, fg_color="#F0F5FA", text_color="#333", hover_color="#DDD").pack(side="left", padx=5)
        ctk.CTkLabel(r_stop, text="(结果标签中剔除这些词)", text_color="gray", font=("Arial", 11)).pack(side="left", padx=5)

        self.frame_stopwords = ctk.CTkScrollableFrame(
            parent, height=40, orientation="horizontal", fg_color="#F9F9F9",
            scrollbar_button_color=SCROLLBAR_COLOR, scrollbar_button_hover_color=SCROLLBAR_HOVER
        )
        self.frame_stopwords.pack(fill="x", pady=(5, 0))
        self.refresh_stopwords_ui()

    def render_run_section(self, parent):
        self.btn_run = ctk.CTkButton(parent, text="🚀 开始 AI 聚类分析", command=self.run_process, height=45, font=("Microsoft YaHei", 16, "bold"), fg_color=THEME_COLOR)
        self.btn_run.pack(fill="x", padx=20, pady=(5, 15))
        ctk.CTkLabel(parent, text="执行日志:", text_color="#333", font=FONT_SUBTITLE).pack(anchor="w", padx=20)
        self.log_box = ctk.CTkTextbox(parent, height=180, fg_color="white", text_color="#333", border_color="#CCC", border_width=1, font=("Consolas", 11))
        self.log_box.pack(fill="x", padx=20, pady=(5, 20))

    # --- 逻辑实现 ---

    def log(self, msg): self.log_box.insert("end", f"> {msg}\n"); self.log_box.see("end")
    def update_granularity_label(self, val):
        v = int(val)
        self.lbl_granularity.configure(text="粗略 (类少)" if v < 30 else "精细 (类多)" if v > 70 else "适中")
        if self.df_clean is not None: self.update_k_recommend_text()

    def refresh_keywords_ui(self):
        for w in self.frame_keywords.winfo_children(): w.destroy()
        if not self.filter_keywords: ctk.CTkLabel(self.frame_keywords, text="暂无过滤词", text_color="gray").pack(padx=5); return
        for kw in self.filter_keywords:
            btn = ctk.CTkButton(self.frame_keywords, text=f"{kw} ×", width=60, height=24, fg_color="#E0E0E0", text_color="#333", hover_color="#D63031", command=lambda k=kw: self.remove_keyword(k))
            btn.pack(side="left", padx=2)

    def add_keyword(self):
        k = self.entry_kw.get().strip()
        if k and k not in self.filter_keywords: self.filter_keywords.append(k); self.entry_kw.delete(0, "end"); self.refresh_keywords_ui()
    def remove_keyword(self, k):
        if k in self.filter_keywords: self.filter_keywords.remove(k); self.refresh_keywords_ui()
    def clear_keywords(self): self.filter_keywords = []; self.refresh_keywords_ui()

    def refresh_stopwords_ui(self):
        for w in self.frame_stopwords.winfo_children(): w.destroy()
        sw_list = sorted(list(self.label_stopwords))
        for sw in sw_list:
            btn = ctk.CTkButton(self.frame_stopwords, text=f"{sw} ×", width=60, height=24, fg_color="#FFF0F0", text_color="#666", hover_color="#FFDEDE", command=lambda k=sw: self.remove_stopword(k))
            btn.pack(side="left", padx=2)
    def add_stopword(self):
        k = self.entry_stop.get().strip()
        if k: self.label_stopwords.add(k); self.entry_stop.delete(0, "end"); self.refresh_stopwords_ui()
    def remove_stopword(self, k):
        if k in self.label_stopwords: self.label_stopwords.remove(k); self.refresh_stopwords_ui()

    def load_file_thread(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls *.csv")])
        if not path: return
        self.btn_file.configure(state="disabled", text="正在读取...")
        self.lbl_file_status.configure(text="读取大文件中...", text_color="#E67E22")
        def task():
            try:
                self.df = pd.read_csv(path) if path.endswith(".csv") else pd.read_excel(path)
                self.file_path = path
                cols = [str(c) for c in self.df.columns.tolist()]
                self.btn_file.after(0, lambda: self._update_ui_after_load(path, cols))
            except Exception as e:
                self.btn_file.after(0, lambda: messagebox.showerror("错误", f"读取失败: {e}"))
                self.btn_file.after(0, lambda: self._reset_load_btn())
        threading.Thread(target=task, daemon=True).start()

    def _reset_load_btn(self):
        self.btn_file.configure(state="normal", text="📂 选择序时账")
        self.lbl_file_status.configure(text="读取失败", text_color="red")

    def _update_ui_after_load(self, path, cols):
        self.btn_file.configure(state="normal", text="📂 重新选择")
        self.lbl_file_status.configure(text=f"已加载: {os.path.basename(path)} ({len(self.df)} 行)", text_color="green")
        for combo in [self.combo_abstract, self.combo_debit, self.combo_credit, self.combo_subject]:
            combo.configure(state="normal", values=cols)
        for c in cols:
            if "摘要" in c or "备注" in c: self.combo_abstract.set(c)
            if "科目" in c and "名称" in c: self.combo_subject.set(c)
            if "借方" in c: self.combo_debit.set(c)
            if "贷方" in c: self.combo_credit.set(c)
        self.log(f"文件加载成功。总行数: {len(self.df)}")

    def clean_text(self, text):
        if not isinstance(text, str): return ""
        if self.var_remove_num.get():
            text = re.sub(r'\d{2,4}[-./]\d{1,2}[-./]\d{1,2}', ' ', text)
            text = re.sub(r'\d{2,4}年|\d{1,2}月|\d{1,2}日', ' ', text)
            text = re.sub(r'\d+\.?\d*', ' ', text)
            text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z]', ' ', text)
        return re.sub(r'\s+', ' ', text).strip()

    def generate_cluster_label(self, texts_in_cluster, group_name):
        # === 【插入】懒加载 jieba ===
        import jieba.analyse 
        # ==========================
        full_text = " ".join(texts_in_cluster)
        keywords = jieba.analyse.extract_tags(full_text, topK=15)
        
        valid_keywords = []
        target_count = int(self.slider_topk.get())
        
        for kw in keywords:
            if kw in self.label_stopwords: continue
            if len(kw) < 2: continue
            valid_keywords.append(kw)
            if len(valid_keywords) >= target_count: break 
        
        kw_str = " ".join(valid_keywords) if valid_keywords else "其他"
        return f"{group_name} | {kw_str}"

    def update_k_recommend_text(self):
        if self.df_clean is None: return
        col_abs = self.combo_abstract.get()
        if col_abs not in self.df_clean.columns: return

        raw_texts = self.df_clean[col_abs].astype(str).fillna("")
        if self.var_remove_num.get():
            cleaned = raw_texts.str.replace(r'\d+', '', regex=True)
            unique_count = cleaned.nunique()
        else:
            unique_count = raw_texts.nunique()
        
        granularity = self.slider_g.get()
        denominator = max(20, 200 - (granularity * 1.8)) 
        rec_k = int(max(2, min(50, unique_count // denominator)))
        
        self.lbl_k_recommend.configure(text=f"💡 当前粒度预计会将每组分为约 {rec_k} 类 (基于数据量)")

    def calculate_stats(self):
        if self.df is None: return messagebox.showwarning("提示", "请先加载文件")
        
        col_abs = self.combo_abstract.get()
        col_jie = self.combo_debit.get()
        col_dai = self.combo_credit.get()
        
        try: threshold = float(self.entry_threshold.get())
        except: return messagebox.showerror("错误", "重要性水平请输入数字")

        mask_keep = pd.Series([True] * len(self.df))
        total_debit = 0.0
        total_credit = 0.0

        if col_jie in self.df.columns or col_dai in self.df.columns:
            jie_vals = pd.to_numeric(self.df[col_jie], errors='coerce').fillna(0) if col_jie in self.df.columns else pd.Series([0]*len(self.df))
            dai_vals = pd.to_numeric(self.df[col_dai], errors='coerce').fillna(0) if col_dai in self.df.columns else pd.Series([0]*len(self.df))
            mask_money = (jie_vals >= threshold) | (dai_vals >= threshold)
            mask_keep = mask_keep & mask_money
        
        if col_abs in self.df.columns and self.filter_keywords:
            pattern = "|".join([re.escape(k) for k in self.filter_keywords])
            has_kw = self.df[col_abs].astype(str).str.contains(pattern, na=False, regex=True)
            mask_keep = mask_keep & (~has_kw)

        if col_jie in self.df.columns:
            jie_vals = pd.to_numeric(self.df[col_jie], errors='coerce').fillna(0)
            total_debit = jie_vals[mask_keep].sum()
        if col_dai in self.df.columns:
            dai_vals = pd.to_numeric(self.df[col_dai], errors='coerce').fillna(0)
            total_credit = dai_vals[mask_keep].sum()

        count_final = mask_keep.sum()
        self.df_clean = self.df[mask_keep].copy()
        
        self.update_k_recommend_text()

        msg = (
            f"原始数据: {len(self.df)} 行\n"
            f"最终保留: {count_final} 行 (过滤率 {100 - (count_final/len(self.df)*100):.1f}%)\n"
            f"剩余金额: 借方 ¥{total_debit:,.2f} | 贷方 ¥{total_credit:,.2f}"
        )
        self.lbl_stats.configure(text=msg)
        return mask_keep

    def run_process(self):
        # === 【修改】把原来的 HAS_NLP 检查换成这个 ===
        try:
            from sentence_transformers import SentenceTransformer
            from sklearn.cluster import KMeans
            import jieba
        except ImportError:
            return messagebox.showerror("环境错误", "未检测到 sentence-transformers 或 sklearn")
        # ==========================================

        # === 【核心修改】Lazy Load，把导入和检查移到这里 ===
        if not self.calculate_stats().any(): return messagebox.showwarning("提示", "有效数据为 0")

        model_path = get_model_path(self.model_name)
        if not model_path: return messagebox.showerror("模型丢失", f"找不到 {self.model_name}")

        col_abs = self.combo_abstract.get()
        col_subject = self.combo_subject.get()
        do_split = self.var_split_subject.get()
        granularity = self.slider_g.get()
        do_clean_num = self.var_remove_num.get()
        
        if do_split and (not col_subject or col_subject not in self.df.columns):
            return messagebox.showwarning("提示", "启用科目分割聚类时，必须选择有效的一级科目列")

        stop_event = None
        if hasattr(self, 'app'): stop_event = self.app.register_task(self.module_index)

        self.btn_run.configure(state="disabled", text="AI 正在思考中...")
        self.log_box.delete("1.0", "end")

        def task():
            try:
                 # === 【插入】在这里加载重型库 ===
                try:
                    from sentence_transformers import SentenceTransformer
                    from sklearn.cluster import KMeans
                    from sklearn.metrics import silhouette_score
                except ImportError:
                    self.log("错误: 未检测到 sentence-transformers 库，无法运行。")
                    return

                self.log(f"加载模型: {self.model_name} ...")
                model = SentenceTransformer(model_path)
                
                df_result = self.df.copy()
                df_result['聚类ID'] = "-1"
                df_result['建议标签'] = ""
                df_result.loc[~df_result.index.isin(self.df_clean.index), '建议标签'] = "(已过滤)"
                
                groups = []
                if do_split:
                    unique_subjects = self.df_clean[col_subject].astype(str).unique()
                    self.log(f"检测到 {len(unique_subjects)} 个科目分组...")
                    for subj in unique_subjects:
                        indices = self.df_clean[self.df_clean[col_subject].astype(str) == subj].index
                        groups.append((subj, indices))
                else:
                    groups.append(("全量", self.df_clean.index))

                total_groups = len(groups)
                
                for g_idx, (g_name, indices) in enumerate(groups):
                    if stop_event and stop_event.is_set(): break
                    
                    raw_texts = self.df_clean.loc[indices, col_abs].astype(str).fillna("").tolist()
                    cleaned_map = {} 
                    for r_idx, txt in zip(indices, raw_texts):
                        c = self.clean_text(txt) if do_clean_num else txt.strip()
                        if len(c) < 1: c = "(杂项)"
                        if c not in cleaned_map: cleaned_map[c] = []
                        cleaned_map[c].append(r_idx)
                    
                    unique_cleaned = list(cleaned_map.keys())
                    n_samples = len(unique_cleaned)
                    
                    if n_samples < 2:
                        for r_idx in indices:
                            df_result.at[r_idx, '聚类ID'] = f"{g_name}_0"
                            df_result.at[r_idx, '建议标签'] = f"{g_name} | 样本过少"
                        continue

                    denominator = max(20, 200 - (granularity * 1.8)) 
                    k = int(n_samples // denominator)
                    k = max(2, min(k, 50))
                    
                    self.log(f"[{g_idx+1}/{total_groups}] 分析 [{g_name}]: {n_samples} 条唯一语义 -> 分 {k} 类")

                    embeddings = model.encode(unique_cleaned, batch_size=64, show_progress_bar=False)
                    kmeans = KMeans(n_clusters=k, random_state=42)
                    labels = kmeans.fit_predict(embeddings)
                    
                    label_names = {}
                    for label_id in range(k):
                        texts_in_cluster = [unique_cleaned[i] for i, x in enumerate(labels) if x == label_id]
                        label_names[label_id] = self.generate_cluster_label(texts_in_cluster, g_name)

                    for u_idx, label_id in enumerate(labels):
                        u_text = unique_cleaned[u_idx]
                        original_indices = cleaned_map[u_text]
                        final_label = label_names[label_id]
                        df_result.loc[original_indices, '聚类ID'] = f"{g_name}_{label_id}"
                        df_result.loc[original_indices, '建议标签'] = final_label

                if not (stop_event and stop_event.is_set()):
                    import time
                    save_path = os.path.splitext(self.file_path)[0] + f"_AI聚类结果_{int(time.time())}.xlsx"
                    df_result.to_excel(save_path, index=False)
                    self.log("-" * 30)
                    self.log(f"结果已保存: {os.path.basename(save_path)}")
                    messagebox.showinfo("成功", "分析完成")
                else:
                    self.log(">>> 任务强制终止")

            except Exception as e:
                # ==================== 将报错写入日志文件 ====================
                import traceback
                import sys
                import os
                
                # 获取日志存放路径（放在 exe 旁边）
                if getattr(sys, 'frozen', False):
                    log_dir = os.path.dirname(sys.executable)
                else:
                    log_dir = os.getcwd()
                
                log_path = os.path.join(log_dir, "error_log.txt")
                
                # 写入详细报错
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write("="*60 + "\n")
                    f.write("CRASH REPORT - 基米工具箱 NLP 模块\n")
                    f.write("="*60 + "\n")
                    f.write(traceback.format_exc())
                    f.write("\n" + "="*60 + "\n")
                
                # 弹窗提示
                messagebox.showerror(
                    "程序崩溃", 
                    f"程序运行时发生错误。\n\n详细的错误日志已保存至:\n{log_path}\n\n请将此文件发给开发者查看。"
                )
                
                self.log_box.insert("end", f"\n发生严重错误，详情见 {log_path}\n")
            finally:
                if hasattr(self, 'app'): self.app.finish_task(self.module_index)
                self.btn_run.configure(state="normal", text="🚀 开始 AI 聚类分析")

        threading.Thread(target=task, daemon=True).start()