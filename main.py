import customtkinter as ctk
import tkinter as tk
from ctypes import windll
import sys
import os
import threading

# === 【修改】删除顶层的模块导入，改为动态导入 ===
# from modules.xls_to_xlsx import XLSToXLSXModule
# from modules.column_extractor import ColumnExtractorModule
# from modules.file_batch_tool import FileBatchToolModule
# from modules.pdf_indexer import PDFIndexerModule
# from modules.pdf_merger import PDFMergerModule
# from modules.id_photo_tool import IDPhotoToolModule
# from modules.sticker_maker import StickerMakerModule
# from modules.keyword_search import keyWordSearchModule
# from modules.smart_extractor import SmartExtractorModule
# from modules.ai_console import AIConsoleModule
# from modules.audit_radar_module import AuditRadarModule
# from modules.nlp_cluster import NLPClusterModule
# from modules.smart_reconciler import SmartReconcilerModule
# from modules.contra_analyzer import ContraAnalyzerModule

# --- 引入图标模块 ---
from modules.path_manager import get_asset_path 

# --- DPI适配 ---
try:
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

ctk.set_appearance_mode("Light")       
ctk.set_default_color_theme("blue")    

# === 【新增】模块注册表（配置哪些模块需要动态导入）===
MODULE_REGISTRY = {
    0: {"name": "AI 控制台", "import_path": "modules.ai_console", "class_name": "AIConsoleModule"},
    1: {"name": "AI 智能文档提取", "import_path": "modules.smart_extractor", "class_name": "SmartExtractorModule"},
    2: {"name": "Excel数据提取", "import_path": "modules.column_extractor", "class_name": "ColumnExtractorModule"},
    3: {"name": "文件批量整理", "import_path": "modules.file_batch_tool", "class_name": "FileBatchToolModule"},
    4: {"name": "内容检索&链接管理", "import_path": "modules.keyword_search", "class_name": "keyWordSearchModule"},
    5: {"name": "Excel 格式互转", "import_path": "modules.xls_to_xlsx", "class_name": "XLSToXLSXModule"},
    6: {"name": "PDF 合并/分拆/转换", "import_path": "modules.pdf_merger", "class_name": "PDFMergerModule"},
    7: {"name": "PDF 索引号生成", "import_path": "modules.pdf_indexer", "class_name": "PDFIndexerModule"},
    8: {"name": "会计分录测试", "import_path": "modules.audit_radar_module", "class_name": "AuditRadarModule"},
    9: {"name": "摘要语义聚类分析", "import_path": "modules.nlp_cluster", "class_name": "NLPClusterModule"},
    10: {"name": "银行流水核查", "import_path": "modules.smart_reconciler", "class_name": "SmartReconcilerModule"},
    11: {"name": "对方科目分析", "import_path": "modules.contra_analyzer", "class_name": "ContraAnalyzerModule"},
    12: {"name": "证件照处理", "import_path": "modules.id_photo_tool", "class_name": "IDPhotoToolModule"},
    13: {"name": "表情包/贴纸生成", "import_path": "modules.sticker_maker", "class_name": "StickerMakerModule"},
}

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- 窗口基础设置 ---
        self.title("基米工具箱 v0.3")
        self.geometry("900x600") 
        self.minsize(900, 600)

        # --- 图标设置 ---
        try:
            icon_path = get_asset_path(os.path.join("assets", "icon.ico"))
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

        # --- 布局容器 (解决卡顿) ---
        self.paned_window = tk.PanedWindow(
            self, 
            orient="horizontal", 
            sashwidth=6, bg="#E5E5E5", bd=0, opaqueresize=False
        )
        self.paned_window.pack(fill="both", expand=True)

        # --- 左侧侧边栏 ---
        self.sidebar_container = ctk.CTkFrame(self.paned_window, corner_radius=0, fg_color="#F3F3F3")
        self.paned_window.add(self.sidebar_container, minsize=280, width=280)
        self.sidebar_container.grid_columnconfigure(0, weight=1)
        self.sidebar_container.grid_rowconfigure(1, weight=1) 

        self.logo_label = ctk.CTkLabel(
            self.sidebar_container, 
            text="工具箱 Pro", 
            font=ctk.CTkFont(family="Microsoft YaHei", size=24, weight="bold"),
            text_color="#333333"
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 15), sticky="w")

        self.nav_scroll = ctk.CTkScrollableFrame(
            self.sidebar_container, fg_color="transparent", corner_radius=0,
            scrollbar_button_color="#E0E0E0", scrollbar_button_hover_color="#D0D0D0"
        )
        self.nav_scroll.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)

        self.version_label = ctk.CTkLabel(
            self.sidebar_container, text="v0.3", font=ctk.CTkFont(family="Arial", size=12), text_color="#999999"
        )
        self.version_label.grid(row=2, column=0, pady=10, sticky="s")

        # --- 右侧内容容器 ---
        self.content_frame = ctk.CTkFrame(self.paned_window, corner_radius=0, fg_color="white")
        self.paned_window.add(self.content_frame, minsize=500, stretch="always")

        # ==================== 架构核心：视图缓存 & 任务管理 ====================
        self.module_frames = {}         # 缓存已渲染的模块视图
        self.module_instances = {}      # 【新增】缓存已加载的模块实例
        self.current_module_index = None
        
        # 任务登记册 { 模块索引: stop_event }
        self.running_tasks = {} 

        # 初始化悬浮球 (默认隐藏)
        self.init_floating_control()

        # ==================== 【修改】模块注册（不再立即实例化）================
        # 原代码：self.modules = [AIConsoleModule(), SmartExtractorModule(), ...]
        # 新代码：改为从注册表动态加载
        
        self.sidebar_buttons = []
        self.init_sidebar()

        if MODULE_REGISTRY:
            self.select_module(0)

    # === 【新增】动态加载模块实例 ===
    def get_module(self, index):
        """
        懒加载模块：第一次使用时才导入并实例化
        返回: 模块实例
        """
        if index not in self.module_instances:
            config = MODULE_REGISTRY[index]
            import importlib
            
            # 动态导入模块
            module = importlib.import_module(config["import_path"])
            
            # 获取类并实例化
            module_class = getattr(module, config["class_name"])
            instance = module_class()
            
            # 缓存实例
            self.module_instances[index] = instance
            print(f"[懒加载] 已加载模块 {index}: {config['name']}")
        
        return self.module_instances[index]

    # ==================== 新增：悬浮中断控制器逻辑 ====================
    def init_floating_control(self):
        """初始化右下角的悬浮控制条"""
        self.float_frame = ctk.CTkFrame(self, fg_color="#333333", corner_radius=10, bg_color="transparent")
        
        self.task_label = ctk.CTkLabel(self.float_frame, text="正在运行: 0", text_color="white", font=("Arial", 12, "bold"))
        self.task_label.pack(side="left", padx=(15, 10), pady=10)

        self.btn_stop_current = ctk.CTkButton(
            self.float_frame, text="停止当前", width=80, height=30,
            fg_color="#FF4D4D", hover_color="#CC0000",
            command=self.stop_current_task
        )
        self.btn_stop_current.pack(side="left", padx=5, pady=10)

        self.btn_stop_all = ctk.CTkButton(
            self.float_frame, text="停止所有", width=80, height=30,
            fg_color="#555555", hover_color="#333333",
            command=self.stop_all_tasks
        )
        self.btn_stop_all.pack(side="left", padx=(5, 15), pady=10)

    def update_floating_visibility(self):
        """根据是否有任务运行，自动显示或隐藏悬浮球"""
        count = len(self.running_tasks)
        if count > 0:
            self.task_label.configure(text=f"后台任务: {count}")
            self.float_frame.place(relx=0.98, rely=0.98, anchor="se")
            self.float_frame.lift()
        else:
            self.float_frame.place_forget()

    # ==================== 新增：对外接口 (供模块调用) ====================
    
    def register_task(self, module_index):
        """
        模块开始运行时调用此方法。
        返回: stop_event (threading.Event)
        """
        stop_event = threading.Event()
        self.running_tasks[module_index] = stop_event
        self.update_floating_visibility()
        return stop_event

    def finish_task(self, module_index):
        """模块运行结束（无论成功失败）调用此方法"""
        self.after(0, lambda: self._finish_task_internal(module_index))

    def _finish_task_internal(self, module_index):
        if module_index in self.running_tasks:
            del self.running_tasks[module_index]
        self.update_floating_visibility()

    # ==================== 新增：按钮响应逻辑 ====================

    def stop_current_task(self):
        """停止当前正在显示的模块的任务"""
        idx = self.current_module_index
        if idx is not None and idx in self.running_tasks:
            print(f"请求停止模块 {idx}...")
            self.running_tasks[idx].set()
            self.btn_stop_current.configure(text="正在停止...", state="disabled")
            self.after(1000, lambda: self.btn_stop_current.configure(text="停止当前", state="normal"))

    def stop_all_tasks(self):
        """停止所有后台任务（带确认机制）"""
        if not self.running_tasks:
            return

        confirm = tk.messagebox.askyesno(
            "高危操作确认", 
            f"当前有 {len(self.running_tasks)} 个任务正在后台运行。\n\n确定要全部强制停止吗？"
        )

        if not confirm:
            return

        print(">>> 用户确认停止所有任务...")
        for idx, event in self.running_tasks.items():
            event.set()
        
        self.btn_stop_all.configure(text="正在停止...", state="disabled")
        self.after(1000, lambda: self.btn_stop_all.configure(text="停止所有", state="normal"))

    # ==================== 原有逻辑（已适配懒加载）================

    def init_sidebar(self):
        """初始化侧边栏按钮"""
        for idx in sorted(MODULE_REGISTRY.keys()):
            config = MODULE_REGISTRY[idx]
            
            btn = ctk.CTkButton(
                self.nav_scroll,  
                text=f"  {config['name']}",  # 从注册表获取名称
                command=lambda i=idx: self.select_module(i),
                fg_color="transparent",
                text_color="#555555",
                hover_color="#E1EBF5", 
                anchor="w",
                height=40, 
                corner_radius=8,
                font=ctk.CTkFont(family="Microsoft YaHei", size=14)
            )
            btn.grid(row=idx, column=0, padx=10, pady=5, sticky="ew")
            self.sidebar_buttons.append(btn)

    def select_module(self, index):
        # 1. 样式更新
        for i, btn in enumerate(self.sidebar_buttons):
            if i == index:
                btn.configure(fg_color="white", text_color="#007AFF") 
            else:
                btn.configure(fg_color="transparent", text_color="#555555")
        
        # 2. 隐藏旧模块
        if self.current_module_index is not None:
            old_frame = self.module_frames.get(self.current_module_index)
            if old_frame:
                old_frame.pack_forget()

        # 3. 显示新模块
        self.current_module_index = index
        
        # === 【修改】懒加载模块 ===
        current_module = self.get_module(index)
        
        # 注入 app 实例给模块
        if not hasattr(current_module, 'app'):
            current_module.app = self
            current_module.module_index = index

        if index in self.module_frames:
            self.module_frames[index].pack(fill="both", expand=True)
        else:
            new_frame = ctk.CTkFrame(self.content_frame, corner_radius=0, fg_color="white")
            new_frame.pack(fill="both", expand=True)
            current_module.render(new_frame)
            self.module_frames[index] = new_frame

if __name__ == "__main__":
    app = App()
    app.mainloop()