import os
import sys
import fitz  # PyMuPDF
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading

# --- 资源路径辅助函数 ---
from modules.path_manager import get_asset_path  # 使用统一的 path_manager，如果没引入，也可以用下面的备用

def get_resource_path_local(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- 核心逻辑 ---

# === 【修改点 1】新增 stop_event 参数 ===
def add_index_to_pdf(file_path, font_path, log_callback, stop_event=None):
    """
    使用 TextWriter 修复中文显示问题，并定位到右上角
    """
    doc = None
    temp_path = file_path + ".tmp" # 提前定义，方便 finally 清理
    
    try:
        try:
            custom_font = fitz.Font(fontfile=font_path)
            if not custom_font.name:
                return False, f"字体加载异常: {font_path}"
        except Exception as e:
            return False, f"加载字体文件失败: {e}"

        doc = fitz.open(file_path)
        filename = os.path.basename(file_path)
        prefix_name = os.path.splitext(filename)[0]
        total_pages = len(doc)
        
        FONT_SIZE = 18          
        MARGIN_RIGHT = 30       
        MARGIN_TOP = 30         
        TEXT_COLOR = (1, 0, 0)  

        # 遍历每一页
        for page_num, page in enumerate(doc):
            # === 【修改点 2】页面级中断检测 ===
            if stop_event and stop_event.is_set():
                doc.close()
                if os.path.exists(temp_path): os.remove(temp_path) # 清理垃圾
                return False, ">>> 用户强制停止（当前文件未修改）"
            # =================================

            text = f"{prefix_name} {page_num + 1}/{total_pages}"
            page_rect = page.rect
            text_length = custom_font.text_length(text, fontsize=FONT_SIZE)
            x = page_rect.width - text_length - MARGIN_RIGHT
            y = MARGIN_TOP + FONT_SIZE  

            tw = fitz.TextWriter(page_rect)
            tw.append(pos=(x, y), text=text, font=custom_font, fontsize=FONT_SIZE)
            tw.write_text(page, color=TEXT_COLOR, render_mode=2)

        doc.save(temp_path)
        doc.close()
        doc = None 

        os.replace(temp_path, file_path)
        return True, f"成功: {filename}"

    except Exception as e:
        import traceback
        if doc: doc.close()
        if os.path.exists(temp_path): os.remove(temp_path)
        return False, f"失败 {os.path.basename(file_path)}: {str(e)}"
    finally:
        if doc: doc.close()

# === 【修改点 3】新增 stop_event 参数 ===
def batch_process_pdf(target_path, log_callback, stop_event=None):
    """批量处理逻辑"""
    # 尝试使用统一的 path_manager，如果没有则用本地备用
    try:
        from modules.path_manager import get_asset_path
        font_path = get_asset_path(os.path.join("assets", "fonts", "simsun.ttc"))
    except ImportError:
        font_path = get_resource_path_local(os.path.join("assets", "fonts", "simsun.ttc"))
    
    log_callback(f"正在加载字体: {font_path}")
    if not os.path.exists(font_path):
        return f"错误: 找不到字体文件!\n请确认文件存在于:\n{font_path}"

    files_to_process = []
    
    if os.path.isfile(target_path):
        if target_path.lower().endswith('.pdf'):
            files_to_process.append(target_path)
    else:
        for root, dirs, files in os.walk(target_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    files_to_process.append(os.path.join(root, file))

    if not files_to_process:
        return "未找到 PDF 文件。"

    log_callback(f"找到 {len(files_to_process)} 个 PDF 文件，准备处理...")
    
    success_count = 0
    fail_count = 0

    for i, file_path in enumerate(files_to_process):
        # === 【修改点 4】文件级中断检测 ===
        if stop_event and stop_event.is_set():
            log_callback(">>> 用户强制停止任务！")
            break
        # ===============================

        # 传递 stop_event 给单个文件处理函数
        status, msg = add_index_to_pdf(file_path, font_path, log_callback, stop_event)
        
        if status:
            success_count += 1
            log_callback(msg)
        else:
            log_callback(msg)
            # 如果是因为用户停止导致的 False，我们就不算作“失败”，而是直接退出
            if "用户强制停止" in msg:
                break
            fail_count += 1

    return f"处理完成。\n成功: {success_count}\n失败: {fail_count}"

# --- 界面模块 ---

class PDFIndexerModule:
    def __init__(self):
        self.name = "PDF 索引号生成"
        self.target_path = ""
        # self.app 会由 main.py 注入

    def render(self, parent_frame):
        for widget in parent_frame.winfo_children(): widget.destroy()

        scroll = ctk.CTkScrollableFrame(parent_frame, fg_color="transparent", scrollbar_button_color="#E0E0E0", scrollbar_button_hover_color="#D0D0D0")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(scroll, text="PDF索引号生成", font=("Microsoft YaHei", 22, "bold"), text_color="#333").pack(pady=15, anchor="w", padx=20)

        tips = """
        【功能说明】
        1. 自动读取 PDF 文件名作为索引前缀。
        2. 格式：文件名 页码/总页数 (例如：XX-1 1/6)。
        3. 位置：右上角 (自动计算宽度对齐)。
        4. 样式：红色、粗体、字号 18。
        5. 注意：操作不可逆，请备份文件。
        """
        tips_frame = ctk.CTkFrame(scroll, fg_color="#F8F9FA", corner_radius=6)
        tips_frame.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(tips_frame, text=tips, justify="left", text_color="#555", font=("Consolas", 12)).pack(padx=15, pady=10, anchor="w")

        op_frame = ctk.CTkFrame(scroll, fg_color="white", corner_radius=8, border_width=1, border_color="#DDD")
        op_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(op_frame, text="选择文件或文件夹:", font=("Microsoft YaHei", 14), text_color="#333").pack(anchor="w", padx=15, pady=(15, 5))
        
        self.entry_path = ctk.CTkEntry(op_frame, placeholder_text="支持单选 PDF 或整个文件夹...", height=35)
        self.entry_path.pack(fill="x", padx=15, pady=5)
        
        btn_box = ctk.CTkFrame(op_frame, fg_color="transparent")
        btn_box.pack(fill="x", padx=10, pady=15)
        
        ctk.CTkButton(btn_box, text="选择文件夹", width=100, command=self.select_folder).pack(side="left", padx=5)
        ctk.CTkButton(btn_box, text="选择单文件", width=100, command=self.select_file).pack(side="left", padx=5)
        
        # 绑定 run_task
        self.btn_run = ctk.CTkButton(btn_box, text="开始添加索引", width=120, fg_color="#d63031", command=self.run_task)
        self.btn_run.pack(side="right", padx=5)

        ctk.CTkLabel(scroll, text="运行日志", text_color="#333", font=("Microsoft YaHei", 14, "bold")).pack(anchor="w", padx=20, pady=(10, 5))
        self.textbox = ctk.CTkTextbox(scroll, height=250, fg_color="white", text_color="#333", border_width=1, border_color="#CCC")
        self.textbox.pack(padx=20, pady=(0, 20), fill="both")

    def select_folder(self):
        p = filedialog.askdirectory()
        if p: self.entry_path.delete(0, "end"); self.entry_path.insert(0, p); self.target_path = p

    def select_file(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if p: self.entry_path.delete(0, "end"); self.entry_path.insert(0, p); self.target_path = p

    def log(self, msg):
        self.textbox.insert("end", msg + "\n")
        self.textbox.see("end")

    # === 【修改点 5】启动任务 ===
    def run_task(self):
        path = self.entry_path.get().strip()
        if not path or not os.path.exists(path): return messagebox.showerror("错误", "路径不存在")
        if not messagebox.askyesno("警告", "此操作将修改原始文件！\n建议先备份数据。\n是否继续？"): return

        # 申请红旗
        stop_event = None
        if hasattr(self, 'app'): stop_event = self.app.register_task(self.module_index)
        
        self.btn_run.configure(state="disabled", text="处理中...")
        self.textbox.delete("1.0", "end")
        
        def t():
            # 传入 stop_event
            result = batch_process_pdf(path, self.log, stop_event=stop_event)
            self.log("-" * 30)
            self.log(result)
            
            # 销假 & 恢复按钮
            if hasattr(self, 'app'): self.app.finish_task(self.module_index)
            self.btn_run.configure(state="normal", text="开始添加索引")
            
            # 只有在非中断状态下才弹成功窗
            if "用户强制停止" not in result:
                messagebox.showinfo("完成", "任务结束")

        threading.Thread(target=t, daemon=True).start()