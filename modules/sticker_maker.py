import os
import sys
import threading
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageFilter, ImageOps, ImageDraw
from modules.path_manager import get_asset_path, get_model_dir_root 

# --- 资源路径辅助 ---
# 优先尝试使用统一路径管理器
# try:
#     from modules.path_manager import get_asset_path
# except ImportError:
#     def get_asset_path(relative_path):
#         if hasattr(sys, '_MEIPASS'):
#             base_path = sys._MEIPASS
#         else:
#             base_path = os.path.abspath(".")
#         return os.path.join(base_path, relative_path)

# def get_resource_path(relative_path):
#     return get_asset_path(relative_path)

# --- 核心算法：加描边 ---

# === 【修改点 1】新增 stop_event 参数 ===
def add_stroke(img_rgba, stroke_width, stroke_color, stop_event=None):
    padding = stroke_width + 10
    new_size = (img_rgba.width + 2 * padding, img_rgba.height + 2 * padding)
    img_padded = Image.new("RGBA", new_size, (0, 0, 0, 0))
    paste_x = padding
    paste_y = padding
    img_padded.paste(img_rgba, (paste_x, paste_y), mask=img_rgba)

    alpha = img_padded.getchannel("A")
    stroke_mask = alpha
    loop_count = stroke_width
    
    # 循环应用滤镜以产生圆润外扩效果
    for _ in range(loop_count):
        # === 【修改点 2】循环内中断检测 ===
        if stop_event and stop_event.is_set():
            return None # 返回空表示中断
        # ===============================
        stroke_mask = stroke_mask.filter(ImageFilter.MaxFilter(3))

    if stroke_color.startswith("#"):
        rgb = tuple(int(stroke_color[i:i+2], 16) for i in (1, 3, 5))
    else:
        rgb = (255, 255, 255)
    
    stroke_layer = Image.new("RGBA", new_size, rgb + (255,))
    final_img = Image.new("RGBA", new_size, (0, 0, 0, 0))
    final_img.paste(stroke_layer, (0, 0), mask=stroke_mask)
    final_img.paste(img_rgba, (paste_x, paste_y), mask=img_rgba)
    
    bbox = final_img.getbbox()
    if bbox:
        final_img = final_img.crop(bbox)
        
    return final_img

# --- 业务逻辑 ---

# === 【修改点 3】新增 stop_event 参数 ===
def process_sticker(src_path, stroke_width, stroke_color, log_callback, stop_event=None):
    # 懒加载
    try:
        import rembg
        import onnxruntime
    except ImportError:
        return None, "错误：未安装 rembg 或 onnxruntime 库"

    try:
        # === 检测点 A ===
        if stop_event and stop_event.is_set(): return None, "用户终止"

        log_callback("加载图像...")
        img = Image.open(src_path).convert("RGBA")
        
        # === 检测点 B ===
        if stop_event and stop_event.is_set(): return None, "用户终止"

        log_callback("AI 正在抠图 (首次需加载模型)...")
        models_root = get_model_dir_root()
        model_path = os.path.join(models_root, "u2net.onnx")
        
        if os.path.exists(model_path):
            os.environ["U2NET_HOME"] = models_root
            session = rembg.new_session(model_name="u2net")
            img_no_bg = rembg.remove(img, session=session)
        else:
            log_callback("下载模型中...")
            img_no_bg = rembg.remove(img)

        # === 检测点 C (AI 算完后立刻检查) ===
        if stop_event and stop_event.is_set(): return None, "用户终止"

        if stroke_width > 0:
            log_callback("正在渲染描边...")
            # 传入 stop_event 到耗时循环中
            result_img = add_stroke(img_no_bg, stroke_width, stroke_color, stop_event)
            
            if result_img is None: # 说明在 add_stroke 内部被掐断了
                return None, "用户终止"
        else:
            result_img = img_no_bg

        return result_img, "完成"

    except Exception as e:
        return None, f"失败: {str(e)}"

# --- 界面模块 ---
class StickerMakerModule:
    def __init__(self):
        self.name = "表情包/贴纸生成"
        self.src_path = None
        self.result_image = None 
        # self.app 会由 main.py 注入

    def render(self, parent_frame):
        # 1. 清空
        for widget in parent_frame.winfo_children():
            widget.destroy()

        # 2. 分隔栏布局
        self.paned_window = tk.PanedWindow(
            parent_frame, 
            orient="horizontal", 
            sashwidth=5, 
            bg="#E5E5E5", 
            bd=0, 
            opaqueresize=False 
        )
        self.paned_window.pack(fill="both", expand=True)

        # ================= 左侧容器 =================
        self.left_container = ctk.CTkFrame(self.paned_window, corner_radius=0, fg_color="#F9F9F9")
        self.paned_window.add(self.left_container, minsize=340, stretch="never")
        
        self.left_scroll = ctk.CTkScrollableFrame(
            self.left_container, 
            fg_color="transparent",
            width=300,
            scrollbar_button_color="#E0E0E0",
            scrollbar_button_hover_color="#D0D0D0",
            corner_radius=0
        )
        self.left_scroll.pack(fill="both", expand=True)

        # --- 标题 ---
        ctk.CTkLabel(self.left_scroll, text="✨ 贴纸工厂", font=("Microsoft YaHei", 22, "bold"), text_color="#333").pack(pady=(20, 5), anchor="w", padx=20)
        ctk.CTkLabel(self.left_scroll, text="AI 智能抠图 + 描边特效", font=("Microsoft YaHei", 12), text_color="#999").pack(anchor="w", padx=20, pady=(0, 15))

        # --- 卡片1：图片 ---
        card1 = ctk.CTkFrame(self.left_scroll, fg_color="white", corner_radius=10)
        card1.pack(fill="x", padx=15, pady=8)
        
        ctk.CTkLabel(card1, text="1. 上传图片", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(anchor="w", padx=15, pady=(15, 10))
        
        self.btn_select = ctk.CTkButton(
            card1, text="点击选择图片...", command=self.select_img, 
            fg_color="#F0F5FF", text_color="#007AFF", hover_color="#E1EBF5", 
            border_width=1, border_color="#007AFF", height=35
        )
        self.btn_select.pack(fill="x", padx=15, pady=(0, 10))

        self.lbl_thumb = ctk.CTkLabel(card1, text="暂无预览", text_color="#CCC", height=100, fg_color="#F8F8F8", corner_radius=6)
        self.lbl_thumb.pack(fill="x", padx=15, pady=(0, 15))

        # --- 卡片2：参数 ---
        card2 = ctk.CTkFrame(self.left_scroll, fg_color="white", corner_radius=10)
        card2.pack(fill="x", padx=15, pady=8)
        
        ctk.CTkLabel(card2, text="2. 效果参数", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(anchor="w", padx=15, pady=(15, 10))

        ctk.CTkLabel(card2, text="描边颜色", text_color="#666", font=("Microsoft YaHei", 12)).pack(anchor="w", padx=15, pady=(5,0))
        self.combo_color = ctk.CTkComboBox(
            card2, 
            values=["#FFFFFF (白色)", "#000000 (黑色)", "#FF0000 (红色)", "#FFD700 (金色)", "#00FF00 (绿色)"],
            height=32, fg_color="white", border_color="#E0E0E0", text_color="#333",
            dropdown_fg_color="white", dropdown_text_color="#333", button_color="#F0F0F0",
            button_hover_color="#E0E0E0", corner_radius=8
        )
        self.combo_color.set("#FFFFFF (白色)")
        self.combo_color.pack(fill="x", padx=15, pady=5)

        width_header = ctk.CTkFrame(card2, fg_color="transparent")
        width_header.pack(fill="x", padx=15, pady=(10, 0))
        
        ctk.CTkLabel(width_header, text="描边粗细", text_color="#666", font=("Microsoft YaHei", 12)).pack(side="left")
        self.lbl_width_val = ctk.CTkLabel(width_header, text="10", text_color="#007AFF", font=("Arial", 12, "bold"))
        self.lbl_width_val.pack(side="right")

        self.slider_width = ctk.CTkSlider(
            card2, from_=0, to=30, number_of_steps=30, 
            command=self.update_width_label, button_color="#007AFF", progress_color="#007AFF"
        )
        self.slider_width.set(10)
        self.slider_width.pack(fill="x", padx=15, pady=(5, 15))

        # --- 运行按钮 ---
        self.btn_run = ctk.CTkButton(
            self.left_scroll, text="✨ 开始制作贴纸", command=self.run_process, 
            fg_color="#00C853", hover_color="#00A844", height=45, 
            corner_radius=22, font=("Microsoft YaHei", 16, "bold")
        )
        self.btn_run.pack(fill="x", padx=20, pady=(20, 40))

        # ================= 右侧预览区 =================
        self.right_frame = ctk.CTkFrame(self.paned_window, fg_color="white", corner_radius=0)
        self.paned_window.add(self.right_frame, stretch="always")

        top_bar = ctk.CTkFrame(self.right_frame, fg_color="transparent", height=50)
        top_bar.pack(fill="x", padx=20, pady=15)
        
        ctk.CTkLabel(top_bar, text="效果预览", font=("Microsoft YaHei", 18, "bold"), text_color="#333").pack(side="left")
        
        self.btn_save = ctk.CTkButton(
            top_bar, text="💾 保存图片", command=self.save_img, 
            state="disabled", fg_color="#007AFF", width=100, corner_radius=8
        )
        self.btn_save.pack(side="right")

        self.preview_frame = ctk.CTkFrame(self.right_frame, fg_color="#F3F3F3", corner_radius=10)
        self.preview_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        self.preview_container = ctk.CTkLabel(
            self.preview_frame, text="← 请在左侧上传图片", 
            text_color="#999", font=("Microsoft YaHei", 14)
        )
        self.preview_container.pack(fill="both", expand=True, padx=10, pady=10)

        self.status_bar = ctk.CTkFrame(self.right_frame, height=30, fg_color="#F9F9F9")
        self.status_bar.pack(fill="x", side="bottom")
        
        self.status_label = ctk.CTkLabel(self.status_bar, text="准备就绪", text_color="#666", font=("Arial", 11), anchor="w")
        self.status_label.pack(side="left", padx=20)

    def update_width_label(self, value):
        self.lbl_width_val.configure(text=f"{int(value)}")

    def select_img(self):
        path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.bmp;*.webp")])
        if path:
            self.src_path = path
            self.btn_select.configure(text="已选择: " + os.path.basename(path))
            
            img = Image.open(path)
            img.thumbnail((260, 150)) 
            ctk_thumb = ctk.CTkImage(img, size=img.size)
            self.lbl_thumb.configure(image=ctk_thumb, text="")
            self.lbl_thumb.image = ctk_thumb 
            self.status_label.configure(text="图片已加载")

    # === 【修改点 4】启动任务 ===
    def run_process(self):
        if not self.src_path: return messagebox.showwarning("提示", "请先上传图片")

        width = int(self.slider_width.get())
        color_str = self.combo_color.get().split(" ")[0]

        # 申请红旗
        stop_event = None
        if hasattr(self, 'app'): stop_event = self.app.register_task(self.module_index)

        self.btn_run.configure(state="disabled", text="AI 计算中...", fg_color="#BBB")
        self.status_label.configure(text="正在进行 AI 抠图与图像合成...")
        
        def task():
            # 传入 stop_event
            res_img, msg = process_sticker(
                self.src_path, width, color_str, 
                self.update_status, 
                stop_event=stop_event
            )
            
            # 销假
            if hasattr(self, 'app'): self.app.finish_task(self.module_index)
            
            if res_img:
                self.result_image = res_img
                display_img = res_img.copy()
                display_img.thumbnail((800, 600))
                ctk_res = ctk.CTkImage(display_img, size=display_img.size)
                
                self.preview_container.configure(image=ctk_res, text="")
                self.preview_container.image = ctk_res
                
                self.btn_save.configure(state="normal")
                self.btn_run.configure(state="normal", text="✨ 再次制作", fg_color="#00C853")
                self.status_label.configure(text="生成成功！")
            else:
                if msg == "用户终止":
                    self.status_label.configure(text="已停止处理")
                else:
                    messagebox.showerror("错误", msg)
                    self.status_label.configure(text="处理失败")
                
                self.btn_run.configure(state="normal", text="✨ 开始制作贴纸", fg_color="#00C853")

        threading.Thread(target=task, daemon=True).start()

    def update_status(self, msg):
        self.status_label.configure(text=msg)

    def save_img(self):
        if not self.result_image: return
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG Image", "*.png")], initialfile="sticker_output.png")
        if path:
            self.result_image.save(path, format="PNG")
            messagebox.showinfo("成功", f"文件已保存")