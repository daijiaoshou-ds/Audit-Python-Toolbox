import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import json
import pandas as pd # 引入 pandas 做导出
from modules.ai_manager import AIManager, TokenManager

# --- 样式常量 ---
COLOR_BG = "#F5F7FA"       
COLOR_CARD = "white"       
COLOR_TEXT_MAIN = "#333333"
COLOR_TEXT_SUB = "#888888"
COLOR_PRIMARY = "#007AFF"
COLOR_SUCCESS = "#00C853"
COLOR_DANGER = "#FF4757"
COLOR_BORDER = "#E0E0E0"

class AIConsoleModule:
    def __init__(self):
        self.name = "AI 控制台"
        self.config = AIManager.load_config()
        self.current_editing_key = None 

    def render(self, parent_frame):
        for w in parent_frame.winfo_children(): w.destroy()

        self.main_scroll = ctk.CTkScrollableFrame(
            parent_frame, 
            fg_color=COLOR_BG,
            scrollbar_button_color="#E0E0E0",  
            scrollbar_button_hover_color="#D0D0D0"
            )
        self.main_scroll.pack(fill="both", expand=True)

        # 标题区
        header = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(25, 15))
        ctk.CTkLabel(header, text="AI 基础设施配置", font=("Microsoft YaHei", 24, "bold"), text_color=COLOR_TEXT_MAIN).pack(side="left")

        # 1. Token 仪表盘
        self.render_dashboard()

        # 2. 编辑/新增区域
        self.render_edit_section()

        # 3. 列表区域
        self.render_list_section()

        # 4. 角色指派
        self.render_role_section()

        # 初始化数据
        self.refresh_list()
        self.refresh_combos()

    # --- Section 1: Dashboard (UI优化+Excel导出) ---
    def render_dashboard(self):
        stats = TokenManager.get_today_stats()
        total_today = sum(stats.values())
        
        card = ctk.CTkFrame(self.main_scroll, fg_color=COLOR_CARD, corner_radius=8, border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="x", padx=30, pady=10)
        
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=15)
        
        # 左侧总数
        left = ctk.CTkFrame(inner, fg_color="transparent")
        left.pack(side="left")
        ctk.CTkLabel(left, text="今日 Token 消耗", font=("Microsoft YaHei", 14), text_color=COLOR_TEXT_SUB).pack(anchor="w")
        
        # 优化显示：如果是0，就显示普通的0，不要蓝色括号
        val_text = f"{total_today:,}" 
        val_color = COLOR_PRIMARY if total_today > 0 else COLOR_TEXT_SUB
        ctk.CTkLabel(left, text=val_text, font=("Impact", 28), text_color=val_color).pack(anchor="w")
        
        # 右侧操作
        right = ctk.CTkFrame(inner, fg_color="transparent")
        right.pack(side="right")
        ctk.CTkButton(right, text="导出消耗明细 (Excel)", width=140, fg_color="transparent", border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN, command=self.export_tokens).pack()

        # 底部简报
        if stats:
            detail_str = " | ".join([f"{k}: {v:,}" for k,v in stats.items()])
            ctk.CTkLabel(card, text=f"明细: {detail_str}", font=("Arial", 11), text_color=COLOR_TEXT_SUB).pack(anchor="w", padx=20, pady=(0, 10))

    # --- Section 2: Edit Form ---
    def render_edit_section(self):
        self.edit_card = ctk.CTkFrame(self.main_scroll, fg_color=COLOR_CARD, corner_radius=8, border_width=1, border_color=COLOR_BORDER)
        self.edit_card.pack(fill="x", padx=30, pady=10)

        header = ctk.CTkFrame(self.edit_card, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=15)
        self.lbl_edit_title = ctk.CTkLabel(header, text="Step 1. 添加/编辑模型", font=("Microsoft YaHei", 15, "bold"), text_color=COLOR_TEXT_MAIN)
        self.lbl_edit_title.pack(side="left")
        
        self.btn_cancel_edit = ctk.CTkButton(header, text="退出编辑", width=80, height=28, fg_color="#EEE", text_color="#666", hover_color="#DDD", command=self.reset_form_to_new)
        
        form = ctk.CTkFrame(self.edit_card, fg_color="transparent")
        form.pack(fill="x", padx=20, pady=(0, 20))
        form.grid_columnconfigure(1, weight=1)
        form.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(form, text="自定义名称:", anchor="w", text_color=COLOR_TEXT_MAIN).grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.entry_name = ctk.CTkEntry(form, height=35, fg_color="#FAFAFA", border_color=COLOR_BORDER, text_color="black")
        self.entry_name.grid(row=0, column=1, padx=10, pady=8, sticky="ew")
        self.entry_name.bind("<KeyRelease>", self.on_form_change)

        ctk.CTkLabel(form, text="模型ID (Model):", anchor="w", text_color=COLOR_TEXT_MAIN).grid(row=0, column=2, padx=10, pady=8, sticky="w")
        self.entry_model = ctk.CTkEntry(form, height=35, fg_color="#FAFAFA", border_color=COLOR_BORDER, text_color="black")
        self.entry_model.grid(row=0, column=3, padx=10, pady=8, sticky="ew")
        self.entry_model.bind("<KeyRelease>", self.on_form_change)

        ctk.CTkLabel(form, text="API Key:", anchor="w", text_color=COLOR_TEXT_MAIN).grid(row=1, column=0, padx=10, pady=8, sticky="w")
        self.entry_key = ctk.CTkEntry(form, height=35, show="*", fg_color="#FAFAFA", border_color=COLOR_BORDER, text_color="black")
        self.entry_key.grid(row=1, column=1, columnspan=3, padx=10, pady=8, sticky="ew")
        self.entry_key.bind("<KeyRelease>", self.on_form_change)

        ctk.CTkLabel(form, text="Base URL:", anchor="w", text_color=COLOR_TEXT_MAIN).grid(row=2, column=0, padx=10, pady=8, sticky="w")
        self.entry_url = ctk.CTkEntry(form, height=35, placeholder_text="OpenAI兼容格式", fg_color="#FAFAFA", border_color=COLOR_BORDER, text_color="black")
        self.entry_url.grid(row=2, column=1, columnspan=3, padx=10, pady=8, sticky="ew")

        btn_row = ctk.CTkFrame(self.edit_card, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0, 20))
        
        self.btn_save = ctk.CTkButton(btn_row, text="💾 保存配置", height=40, fg_color=COLOR_BORDER, text_color="#999", state="disabled", command=self.save_provider)
        self.btn_save.pack(side="right")

    # --- Section 3: List ---
    def render_list_section(self):
        list_frame = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        list_frame.pack(fill="x", padx=30, pady=5)
        
        ctk.CTkLabel(list_frame, text="Step 2. 已添加的模型", font=("Microsoft YaHei", 15, "bold"), text_color=COLOR_TEXT_MAIN).pack(anchor="w", pady=5)
        
        # 列表高度调小一点
        self.list_container = ctk.CTkScrollableFrame(
            list_frame, 
            height=160, 
            fg_color="white", 
            corner_radius=8, 
            border_width=1, 
            border_color=COLOR_BORDER,
            scrollbar_button_color="#E0E0E0",  
            scrollbar_button_hover_color="#D0D0D0"
            )
        self.list_container.pack(fill="x")

    # --- Section 4: Roles (对齐优化) ---
    def render_role_section(self):
        card = ctk.CTkFrame(self.main_scroll, fg_color=COLOR_CARD, corner_radius=8, border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="x", padx=30, pady=20)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=15)
        ctk.CTkLabel(header, text="Step 3. 引擎指派", font=("Microsoft YaHei", 15, "bold"), text_color=COLOR_TEXT_MAIN).pack(side="left")

        # 网格对齐
        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(fill="x", padx=20, pady=(0, 20))
        
        # 定义列宽：Icon(窄) | Text(定宽) | Combo(拉伸) | Desc(自适应)
        grid.grid_columnconfigure(2, weight=1) 

        # ComboBox 样式
        combo_style = {"height": 35, "fg_color": "white", "button_color": "#F0F0F0", "button_hover_color": "#E0E0E0", "border_color": COLOR_BORDER, "text_color": "black", "dropdown_fg_color": "white", "dropdown_text_color": "black"}

        # --- 视觉引擎 ---
        # 拆分 Icon 和 文字，强制紧挨
        ctk.CTkLabel(grid, text="👁️", width=30).grid(row=0, column=0, pady=10, sticky="e")
        ctk.CTkLabel(grid, text="视觉引擎:", font=("Microsoft YaHei", 13), text_color=COLOR_TEXT_MAIN, anchor="w").grid(row=0, column=1, pady=10, padx=(0, 10), sticky="w")
        
        self.combo_eye = ctk.CTkComboBox(grid, command=self.on_role_change, **combo_style)
        self.combo_eye.grid(row=0, column=2, pady=10, sticky="ew")
        
        ctk.CTkLabel(grid, text="(负责 OCR、截图识别)", text_color=COLOR_TEXT_SUB, font=("Arial", 11), anchor="w").grid(row=0, column=3, padx=10, sticky="w")

        # --- 思考引擎 ---
        ctk.CTkLabel(grid, text="🧠", width=30).grid(row=1, column=0, pady=10, sticky="e")
        ctk.CTkLabel(grid, text="思考引擎:", font=("Microsoft YaHei", 13), text_color=COLOR_TEXT_MAIN, anchor="w").grid(row=1, column=1, pady=10, padx=(0, 10), sticky="w")
        
        self.combo_brain = ctk.CTkComboBox(grid, command=self.on_role_change, **combo_style)
        self.combo_brain.grid(row=1, column=2, pady=10, sticky="ew")
        
        ctk.CTkLabel(grid, text="(负责 提取、清洗、逻辑)", text_color=COLOR_TEXT_SUB, font=("Arial", 11), anchor="w").grid(row=1, column=3, padx=10, sticky="w")

        # --- 推理开关 ---
        # 放在第2列对齐
        self.var_think_mode = ctk.BooleanVar(value=self.config["roles"].get("enable_think", False))
        self.chk_think = ctk.CTkCheckBox(grid, text="启用推理模式 (仅限 r1 等推理模型)", variable=self.var_think_mode, command=self.on_role_change, text_color=COLOR_TEXT_SUB, font=("Arial", 12))
        self.chk_think.grid(row=2, column=2, pady=5, sticky="w")

    # ================= 逻辑 =================

    def on_form_change(self, event=None):
        if self.entry_name.get() and self.entry_key.get() and self.entry_model.get():
            self.btn_save.configure(state="normal", fg_color=COLOR_SUCCESS, text_color="white")
        else:
            self.btn_save.configure(state="disabled", fg_color=COLOR_BORDER, text_color="#999")

    def refresh_list(self):
        for w in self.list_container.winfo_children(): w.destroy()
        
        providers = self.config["providers"]
        if not providers:
            ctk.CTkLabel(self.list_container, text="列表为空", text_color="#CCC").pack(pady=20)
            return

        for name, info in providers.items():
            row = ctk.CTkFrame(self.list_container, fg_color="transparent", height=40)
            row.pack(fill="x", pady=2, padx=5)
            
            bg = "#E3F2FD" if name == self.current_editing_key else "#F9F9F9"
            card = ctk.CTkFrame(row, fg_color=bg, corner_radius=6)
            card.pack(fill="both", expand=True)
            
            def edit_cmd(n=name): self.load_to_form(n)

            ctk.CTkLabel(card, text=name, font=("Microsoft YaHei", 12, "bold"), text_color=COLOR_TEXT_MAIN).pack(side="left", padx=10, pady=8)
            ctk.CTkLabel(card, text=f"({info['model']})", font=("Arial", 11), text_color=COLOR_TEXT_SUB).pack(side="left", pady=8)

            btn_box = ctk.CTkFrame(card, fg_color="transparent")
            btn_box.pack(side="right", padx=5)

            ctk.CTkButton(
                btn_box, text="删除", width=50, height=24, 
                fg_color="#FFEBEE", text_color=COLOR_DANGER, hover_color="#FFCDD2",
                command=lambda n=name: self.delete_provider(n)
            ).pack(side="right", padx=5)

            ctk.CTkButton(
                btn_box, text="编辑", width=50, height=24, 
                fg_color="white", text_color=COLOR_PRIMARY, border_width=1, border_color=COLOR_PRIMARY,
                command=edit_cmd
            ).pack(side="right")

    def load_to_form(self, name):
        self.current_editing_key = name
        info = self.config["providers"][name]

        self.entry_name.delete(0, "end"); self.entry_name.insert(0, name)
        self.entry_url.delete(0, "end"); self.entry_url.insert(0, info.get("url", ""))
        self.entry_key.delete(0, "end"); self.entry_key.insert(0, info["key"])
        self.entry_model.delete(0, "end"); self.entry_model.insert(0, info["model"])

        self.lbl_edit_title.configure(text=f"📝 编辑中: {name}", text_color=COLOR_PRIMARY)
        self.btn_cancel_edit.pack(side="right", padx=10) 
        self.on_form_change() 
        self.refresh_list() 

    def reset_form_to_new(self):
        self.current_editing_key = None
        self.entry_name.delete(0, "end")
        self.entry_url.delete(0, "end")
        self.entry_key.delete(0, "end")
        self.entry_model.delete(0, "end")
        
        self.lbl_edit_title.configure(text="Step 1. 添加/编辑模型", text_color=COLOR_TEXT_MAIN)
        self.btn_cancel_edit.pack_forget()
        self.btn_save.configure(state="disabled", fg_color=COLOR_BORDER)
        self.refresh_list()

    def save_provider(self):
        name = self.entry_name.get().strip()
        url = self.entry_url.get().strip()
        key = self.entry_key.get().strip()
        model = self.entry_model.get().strip()

        if not name or not url or not key:
            messagebox.showwarning("提示", "名称、URL、Key 均为必填项")
            return

        if self.current_editing_key and self.current_editing_key != name:
            del self.config["providers"][self.current_editing_key]

        self.config["providers"][name] = {"url": url, "key": key, "model": model}
        AIManager.save_config(self.config)
        
        self.reset_form_to_new()
        self.refresh_combos()
        messagebox.showinfo("成功", "保存成功")

    def delete_provider(self, name):
        if messagebox.askyesno("确认", f"删除 [{name}] ?"):
            del self.config["providers"][name]
            if self.config["roles"]["vision"] == name: self.config["roles"]["vision"] = None
            if self.config["roles"]["brain"] == name: self.config["roles"]["brain"] = None
            
            AIManager.save_config(self.config)
            
            if self.current_editing_key == name:
                self.reset_form_to_new()
            
            self.refresh_list()
            self.refresh_combos()

    def refresh_combos(self):
        opts = ["(不使用)"] + list(self.config["providers"].keys())
        self.combo_eye.configure(values=opts)
        self.combo_brain.configure(values=opts)
        
        eye = self.config["roles"].get("vision", "(不使用)")
        brain = self.config["roles"].get("brain", "(不使用)")
        
        self.combo_eye.set(eye if eye in opts else "(不使用)")
        self.combo_brain.set(brain if brain in opts else "(不使用)")

    def on_role_change(self, _=None):
        eye = self.combo_eye.get()
        brain = self.combo_brain.get()
        
        self.config["roles"]["vision"] = None if eye == "(不使用)" else eye
        self.config["roles"]["brain"] = None if brain == "(不使用)" else brain
        self.config["roles"]["enable_think"] = self.var_think_mode.get()
        
        AIManager.save_config(self.config)

    def export_tokens(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="Token统计.xlsx")
        if save_path:
            # 读取原始 JSON
            try:
                with open("token_usage.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                # 转为 DataFrame
                rows = []
                for date, models in data.items():
                    for model, count in models.items():
                        rows.append({"日期": date, "模型": model, "消耗Token": count})
                
                if rows:
                    pd.DataFrame(rows).to_excel(save_path, index=False)
                    messagebox.showinfo("成功", "已导出为 Excel")
                else:
                    messagebox.showwarning("提示", "暂无 Token 记录")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")