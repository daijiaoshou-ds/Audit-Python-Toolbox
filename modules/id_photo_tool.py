import os
import io
import threading
import sys
from PIL import Image, ImageFile
import customtkinter as ctk
from tkinter import filedialog, messagebox

# --- 【关键修正】只保留这一行导入，不要在下面自己写函数 ---
from modules.path_manager import get_asset_path, get_model_dir_root

# 允许截断图片加载
ImageFile.LOAD_TRUNCATED_IMAGES = True

# --- 配置 ---
SIZES_CONFIG = {
    "不修改尺寸": None,
    "一寸 (25x35mm)": (295, 413),
    "小一寸 (22x32mm)": (260, 378),
    "二寸 (35x49mm)": (413, 579),
    "小二寸 (33x48mm)": (390, 567),
    "大二寸 (35x53mm)": (413, 626),
    "五寸 (89x127mm)": (1051, 1500),
}

COLORS_CONFIG = {
    "不修改底色": None,
    "白色": (255, 255, 255),
    "红色": (255, 0, 0),
    "蓝色": (67, 142, 219),
    "透明 (PNG)": "TRANSPARENT"
}

# --- 核心处理逻辑 ---
def process_single_image(src_path, save_dir, size_mode, custom_w, custom_h, color_mode, compress_kb, log_callback):
    try:
        filename = os.path.basename(src_path)
        img = Image.open(src_path).convert("RGBA")

        # 1. AI 换底
        if color_mode and color_mode != "不修改底色":
            log_callback(f"  > 正在加载 AI 模型并抠图...")
            try:
                import rembg
                import onnxruntime
            except ImportError:
                return False, "错误: 缺少 AI 库，无法换底色"

            # === 使用统一管理的外部路径 ===
            models_root = get_model_dir_root() 
            model_path = os.path.join(models_root, "u2net.onnx")
            
            if os.path.exists(model_path):
                os.environ["U2NET_HOME"] = models_root
                session = rembg.new_session(model_name="u2net") 
                img_no_bg = rembg.remove(img, session=session)
            else:
                log_callback(f"  [警告] 未找到离线模型({model_path})，尝试联网下载...")
                img_no_bg = rembg.remove(img)

            target_color = COLORS_CONFIG[color_mode]
            if target_color == "TRANSPARENT":
                img = img_no_bg
            else:
                bg_img = Image.new("RGBA", img_no_bg.size, target_color + (255,))
                bg_img.paste(img_no_bg, (0, 0), img_no_bg)
                img = bg_img.convert("RGB")

        # 2. 修改尺寸
        target_w, target_h = 0, 0
        if size_mode == "自定义":
            if custom_w > 0 and custom_h > 0:
                target_w = int(custom_w * 300 / 25.4)
                target_h = int(custom_h * 300 / 25.4)
        elif size_mode in SIZES_CONFIG and SIZES_CONFIG[size_mode]:
            target_w, target_h = SIZES_CONFIG[size_mode]
        
        if target_w > 0 and target_h > 0:
            current_w, current_h = img.size
            target_ratio = target_w / target_h
            current_ratio = current_w / current_h
            
            if current_ratio > target_ratio:
                new_w = int(current_h * target_ratio)
                offset = (current_w - new_w) // 2
                img = img.crop((offset, 0, offset + new_w, current_h))
            else:
                new_h = int(current_w / target_ratio)
                offset = (current_h - new_h) // 2
                img = img.crop((0, offset, current_w, offset + new_h))
            
            img = img.resize((target_w, target_h), Image.Resampling.LANCZOS)

        # 3. 保存与压缩
        target_name = f"pro_{filename}"
        if color_mode == "透明 (PNG)":
            target_name = os.path.splitext(target_name)[0] + ".png"
            final_path = os.path.join(save_dir, target_name)
            img.save(final_path, format="PNG")
            return True, f"成功: {target_name}"

        if img.mode == "RGBA": img = img.convert("RGB")
        target_name = os.path.splitext(target_name)[0] + ".jpg"
        final_path = os.path.join(save_dir, target_name)

        if compress_kb > 0:
            target_bytes = compress_kb * 1024
            quality = 95
            while quality >= 10:
                buffer = io.BytesIO()
                img.save(buffer, format="JPEG", quality=quality)
                if buffer.tell() <= target_bytes:
                    with open(final_path, "wb") as f: f.write(buffer.getvalue())
                    return True, f"成功: {target_name} ({int(buffer.tell()/1024)}KB)"
                quality -= 5
            img.save(final_path, format="JPEG", quality=10)
            return True, f"警告: 压缩极限已达，保存为最低质量"
        else:
            img.save(final_path, format="JPEG", quality=95)
            return True, f"成功: {target_name}"

    except Exception as e:
        return False, f"出错 {filename}: {str(e)}"

# === 批量处理 ===
def batch_process_photos(input_paths, save_dir, size_mode, cw, ch, color, compress, log_callback, stop_event=None):
    success = 0
    fail = 0
    if not os.path.exists(save_dir): os.makedirs(save_dir)
    
    for i, path in enumerate(input_paths):
        # 中断检测
        if stop_event and stop_event.is_set():
            log_callback(">>> 用户强制停止任务！")
            return f"任务已强制终止。\n成功: {success}, 失败: {fail}"

        log_callback(f"[{i+1}/{len(input_paths)}] 处理: {os.path.basename(path)}")
        status, msg = process_single_image(path, save_dir, size_mode, cw, ch, color, compress, log_callback)
        log_callback(msg)
        if status: success += 1
        else: fail += 1
    
    return f"完成。成功: {success}, 失败: {fail}"

# --- 界面模块 ---
class IDPhotoToolModule:
    def __init__(self):
        self.name = "证件照处理"
        self.input_paths = []
        self.save_dir = ""

    def render(self, parent_frame):
        for widget in parent_frame.winfo_children():
            widget.destroy()

        scroll = ctk.CTkScrollableFrame(
            parent_frame, 
            fg_color="transparent",
            scrollbar_button_color="#E0E0E0",  
            scrollbar_button_hover_color="#D0D0D0"
        )
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            scroll, 
            text="证件照处理 (AI换底/改尺寸/压缩)", 
            font=("Microsoft YaHei", 24, "bold"), 
            text_color="#333"
        ).pack(pady=(10, 15), anchor="w", padx=20)

        # ==================== Step 1: 文件选择 ====================
        step1_frame = ctk.CTkFrame(scroll, fg_color="white", corner_radius=10, border_color="#E5E5E5", border_width=1)
        step1_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        ctk.CTkLabel(step1_frame, text="Step 1  选择照片", font=("Microsoft YaHei", 15, "bold"), text_color="#0984e3").pack(anchor="w", padx=20, pady=(10, 5))

        content_f1 = ctk.CTkFrame(step1_frame, fg_color="transparent")
        content_f1.pack(fill="x", padx=20, pady=(0, 15))
        
        ctk.CTkButton(content_f1, text="选择文件夹 (批量)", width=130, height=34, fg_color="#F0F5FA", text_color="#0984e3", hover_color="#E1EBF5", command=self.select_folder).pack(side="left", padx=(0, 10))
        ctk.CTkButton(content_f1, text="选择单张/多张", width=130, height=34, fg_color="#F0F5FA", text_color="#0984e3", hover_color="#E1EBF5", command=self.select_files).pack(side="left")

        self.lbl_count = ctk.CTkLabel(content_f1, text="未选择文件", text_color="#999", font=("Microsoft YaHei", 13))
        self.lbl_count.pack(side="left", padx=15)

        # ==================== Step 2: 参数设置 ====================
        step2_frame = ctk.CTkFrame(scroll, fg_color="white", corner_radius=10, border_color="#E5E5E5", border_width=1)
        step2_frame.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(step2_frame, text="Step 2  参数设置", font=("Microsoft YaHei", 15, "bold"), text_color="#0984e3").pack(anchor="w", padx=20, pady=(10, 5))

        grid_frame = ctk.CTkFrame(step2_frame, fg_color="transparent")
        grid_frame.pack(fill="x", padx=20, pady=(0, 15))
        grid_frame.grid_columnconfigure(0, weight=0, minsize=100)
        grid_frame.grid_columnconfigure(1, weight=1)

        label_style = {"font": ("Microsoft YaHei", 14), "text_color": "#333", "anchor": "e"}
        combo_style = {"width": 240, "height": 34, "fg_color": "#FAFAFA", "border_color": "#D0D0D0", "text_color": "#333333", "button_color": "#0984e3", "button_hover_color": "#076BB8", "dropdown_fg_color": "white", "dropdown_text_color": "#333", "font": ("Microsoft YaHei", 13), "corner_radius": 6}

        # --- Row 1 ---
        ctk.CTkLabel(grid_frame, text="目标尺寸：", **label_style).grid(row=0, column=0, sticky="e", padx=(0, 15), pady=4)
        size_box = ctk.CTkFrame(grid_frame, fg_color="transparent")
        size_box.grid(row=0, column=1, sticky="w", pady=4)
        self.combo_size = ctk.CTkComboBox(size_box, values=list(SIZES_CONFIG.keys()) + ["自定义"], command=self.on_size_change, **combo_style)
        self.combo_size.set("不修改尺寸")
        self.combo_size.pack(side="left")

        self.frame_custom = ctk.CTkFrame(size_box, fg_color="transparent")
        entry_style = {"width": 60, "height": 32, "fg_color": "#FAFAFA", "border_color": "#D0D0D0", "text_color": "#333", "corner_radius": 6}
        self.entry_w = ctk.CTkEntry(self.frame_custom, placeholder_text="W", **entry_style)
        self.entry_w.pack(side="left", padx=(10, 5))
        ctk.CTkLabel(self.frame_custom, text="x", text_color="#999").pack(side="left")
        self.entry_h = ctk.CTkEntry(self.frame_custom, placeholder_text="H", **entry_style)
        self.entry_h.pack(side="left", padx=(5, 5))
        ctk.CTkLabel(self.frame_custom, text="mm", text_color="#999").pack(side="left")

        # --- Row 2 ---
        ctk.CTkLabel(grid_frame, text="背景底色：", **label_style).grid(row=1, column=0, sticky="e", padx=(0, 15), pady=4)
        color_box = ctk.CTkFrame(grid_frame, fg_color="transparent")
        color_box.grid(row=1, column=1, sticky="w", pady=4)
        self.combo_color = ctk.CTkComboBox(color_box, values=list(COLORS_CONFIG.keys()), **combo_style)
        self.combo_color.set("不修改底色")
        self.combo_color.pack(side="left")
        ctk.CTkLabel(color_box, text="(AI 自动抠图)", text_color="#999", font=("Microsoft YaHei", 12)).pack(side="left", padx=10)

        # --- Row 3 ---
        ctk.CTkLabel(grid_frame, text="文件压缩：", **label_style).grid(row=2, column=0, sticky="e", padx=(0, 15), pady=4)
        compress_box = ctk.CTkFrame(grid_frame, fg_color="transparent")
        compress_box.grid(row=2, column=1, sticky="w", pady=4)
        self.entry_kb = ctk.CTkEntry(compress_box, placeholder_text="留空不压缩", width=140, height=34, fg_color="#FAFAFA", border_color="#D0D0D0", text_color="#333", corner_radius=6)
        self.entry_kb.pack(side="left")
        ctk.CTkLabel(compress_box, text="KB (限制最大体积)", text_color="#999").pack(side="left", padx=10)

        # ==================== Start Button ====================
        self.btn_run = ctk.CTkButton(scroll, text="开始处理", height=45, fg_color="#d63031", hover_color="#C02829", font=("Microsoft YaHei", 18, "bold"), corner_radius=8, command=self.run_task)
        self.btn_run.pack(fill="x", padx=20, pady=(10, 15))

        # ==================== Log ====================
        ctk.CTkLabel(scroll, text="运行日志", anchor="w", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(padx=20, anchor="w", pady=(0, 5))
        self.textbox = ctk.CTkTextbox(scroll, height=150, fg_color="#FAFAFA", text_color="#333", border_color="#D0D0D0", border_width=1, corner_radius=6)
        self.textbox.pack(padx=20, pady=(0, 20), fill="both")

    def on_size_change(self, choice):
        if choice == "自定义":
            self.frame_custom.pack(side="left", padx=10)
        else:
            self.frame_custom.pack_forget()

    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Images", "*.jpg *.jpeg *.png")])
        if paths:
            self.input_paths = list(paths)
            self.lbl_count.configure(text=f"已选择 {len(paths)} 张照片", text_color="#0984e3")
            self.save_dir = os.path.dirname(paths[0])

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            files = []
            for f in os.listdir(folder):
                if f.lower().endswith(('.jpg', '.jpeg', '.png')):
                    files.append(os.path.join(folder, f))
            self.input_paths = files
            self.lbl_count.configure(text=f"文件夹内包含 {len(files)} 张图片", text_color="#0984e3")
            self.save_dir = folder

    def log(self, msg):
        self.textbox.insert("end", msg + "\n")
        self.textbox.see("end")

    def run_task(self):
        if not self.input_paths: return messagebox.showwarning("提示", "请先选择照片")

        size_mode = self.combo_size.get()
        color_mode = self.combo_color.get()
        
        cw, ch = 0, 0
        if size_mode == "自定义":
            try:
                cw = float(self.entry_w.get())
                ch = float(self.entry_h.get())
            except:
                return messagebox.showerror("错误", "自定义尺寸请输入有效的数字")
        
        compress_kb = 0
        kb_str = self.entry_kb.get().strip()
        if kb_str:
            try: compress_kb = int(kb_str)
            except: return messagebox.showerror("错误", "压缩大小请输入整数 (KB)")

        if not os.path.exists(os.path.join(self.save_dir, "output")):
            os.makedirs(os.path.join(self.save_dir, "output"))
        final_save_dir = os.path.join(self.save_dir, "output")

        # 申请红旗
        stop_event = None
        if hasattr(self, 'app'): stop_event = self.app.register_task(self.module_index)

        self.btn_run.configure(state="disabled", text="正在处理 (AI运算较慢)...")
        self.textbox.delete("1.0", "end")
        
        def t():
            msg = batch_process_photos(
                self.input_paths, final_save_dir, size_mode, cw, ch, color_mode, compress_kb, self.log,
                stop_event=stop_event
            )
            self.log("-" * 30)
            self.log(msg)
            
            if hasattr(self, 'app'): self.app.finish_task(self.module_index)
            self.btn_run.configure(state="normal", text="开始处理")
            
            if "用户强制停止" not in msg:
                self.log(f"文件保存至: {final_save_dir}")
                messagebox.showinfo("完成", "处理结束")

        threading.Thread(target=t, daemon=True).start()