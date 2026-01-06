import os
import sys
import json
import threading
import time
import base64
import concurrent.futures
import pandas as pd
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image
from pydantic import create_model, Field, ValidationError
from typing import List, Optional
from modules.ai_manager import AIManager, TokenManager
from modules.path_manager import get_schema_dir
# import fitz  # PyMuPDF

# 默认用户指令 (保持不变)
DEFAULT_USER_INSTRUCTION = """这是一份文档。
请帮我提取关键信息。
"""

# --- LogicCore 和 process_pipeline_task 保持完全不变 ---
# (为了节省篇幅，这里折叠 LogicCore 和 process_pipeline_task 的代码，
# 请直接保留你原文件里的这两部分，一字不改)
class LogicCore:
    # ... (保持原代码不变) ...
    @staticmethod
    def load_schemas():
        schemas = {}
        schema_dir = get_schema_dir()
        for f in os.listdir(schema_dir):
            if f.endswith(".json"):
                try:
                    with open(os.path.join(schema_dir, f), "r", encoding="utf-8") as file:
                        data = json.load(file)
                        schemas[data["name"]] = data
                except: pass
        return schemas

    @staticmethod
    def save_schema(name, fields, user_instruction, temperature):
        data = {
            "name": name, 
            "fields": fields,
            "instruction": user_instruction,
            "temperature": temperature
        }
        path = os.path.join(get_schema_dir(), f"{name}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return path

    @staticmethod
    def delete_schema(name):
        path = os.path.join(get_schema_dir(), f"{name}.json")
        if os.path.exists(path):
            os.remove(path)
            return True
        return False

    @staticmethod
    def create_dynamic_model(fields):
        field_definitions = {}
        for f in fields:
            t = str
            if f['type'] == '数字(小数)': t = float
            elif f['type'] == '数字(整数)': t = int
            field_definitions[f['name']] = (Optional[t], Field(None))
        return create_model('DynamicModel', **field_definitions)

    @staticmethod
    def build_final_prompt(user_instruction, ocr_context, fields=None):
        prompt = f"""
### 原始内容 (OCR识别结果)
{ocr_context[:60000]}

### 任务要求
{user_instruction}
"""
        if fields:
            field_desc = "\n".join([f"- Key: \"{f['name']}\" (类型: {f['type']})" for f in fields])
            prompt += f"""
### 输出格式强制要求
1. 你必须且只能输出标准的 JSON 数组 (List[Object])。
2. 严格提取以下字段作为 Key：
{field_desc}
3. 如果没有有效数据，输出 []。
4. 不要输出 Markdown 标记 (```json)，只输出纯文本 JSON。
"""
        return prompt

    @staticmethod
    def pdf_to_base64_images(path):

        import fitz

        images = []
        doc = fitz.open(path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=200) 
            img_data = pix.tobytes("jpeg")
            b64 = base64.b64encode(img_data).decode("utf-8")
            images.append(f"data:image/jpeg;base64,{b64}")
        return images
    
    @staticmethod
    def pdf_to_text(path):

        import fitz

        texts = []
        doc = fitz.open(path)
        for page in doc:
            texts.append(page.get_text())
        return texts

def process_pipeline_task(
    input_data,          
    is_image_input,      
    use_eye,             
    use_brain,           
    current_schema,      
    user_instruction, 
    user_temperature,     
    client_eye, eye_model,
    client_brain, brain_model,
    save_raw_log=False
):
    # ... (保持原代码不变) ...
    raw_text_content = ""
    brain_raw_response = ""
    token_usage = {"eye": 0, "brain": 0}

    # --- Stage 1: Eye-AI ---
    eye_executed = False
    
    if use_eye and client_eye and is_image_input:
        try:
            response = client_eye.chat.completions.create(
                model=eye_model,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": "OCR all text in this image."},
                            {"type": "image_url", "image_url": {"url": input_data}}
                        ]
                    }
                ],
                temperature=0.1
            )
            raw_text_content = response.choices[0].message.content.strip()
            eye_executed = True
            
            if hasattr(response, 'usage'):
                token_usage['eye'] = response.usage.total_tokens
                TokenManager.log_usage(eye_model, response.usage.total_tokens)

        except Exception as e:
            return None, "", f"[视觉引擎错误] {str(e)}", token_usage
    
    elif not is_image_input:
        raw_text_content = input_data
    
    # 如果只开 Eye
    if not use_brain:
        return None, raw_text_content, None, token_usage

    # --- Stage 2: Brain-AI ---
    if not client_brain: return None, "", "思考引擎未配置", token_usage

    context_str = raw_text_content if raw_text_content else "(见附图)"
    final_prompt = LogicCore.build_final_prompt(user_instruction, context_str, current_schema)

    brain_messages = []
    if raw_text_content:
        brain_messages = [{"role": "user", "content": final_prompt}]
    elif is_image_input and use_brain:
        brain_messages = [
            {
                "role": "user", 
                "content": [
                    {"type": "text", "text": final_prompt},
                    {"type": "image_url", "image_url": {"url": input_data}}
                ]
            }
        ]
    else:
        return None, "", "无有效输入", token_usage

    try:
        response = client_brain.chat.completions.create(
            model=brain_model,
            messages=brain_messages,
            temperature=user_temperature
        )
        brain_response_text = response.choices[0].message.content.strip()
        
        if hasattr(response, 'usage'):
            token_usage['brain'] = response.usage.total_tokens
            TokenManager.log_usage(brain_model, response.usage.total_tokens)
        
        # === 分流处理 ===
        if not current_schema:
            return None, brain_response_text, None, token_usage

        content = brain_response_text.replace("```json", "").replace("```", "").strip()
        start = content.find('[')
        end = content.rfind(']')
        if start != -1 and end != -1:
            content = content[start:end+1]
        
        try:
            raw_data = json.loads(content)
        except:
            return None, brain_response_text, f"[JSON解析失败] 见输出文件", token_usage

        if isinstance(raw_data, dict): raw_data = [raw_data]
        
        model_class = LogicCore.create_dynamic_model(current_schema)
        valid_rows = []
        for item in raw_data:
            try:
                obj = model_class(**item)
                valid_rows.append(obj.model_dump())
            except ValidationError:
                continue
        
        return valid_rows, brain_response_text, None, token_usage

    except Exception as e:
        return None, "", f"[思考引擎错误] {str(e)}", token_usage


# ================= 界面模块 =================

class SmartExtractorModule:
    def __init__(self):
        self.name = "AI 智能文档提取"
        self.schemas = LogicCore.load_schemas()
        self.current_schema_fields = [] 
        self.src_path = None
        # === 【修改点 1】删除 self.stop_event 初始化 ===
        # self.stop_event = threading.Event() 
        # 我们将在运行时从 main.py 获取 stop_event

    def render(self, parent_frame):
        for w in parent_frame.winfo_children(): w.destroy()
        
        main_scroll = ctk.CTkScrollableFrame(
            parent_frame, 
            fg_color="#F2F4F8",
            scrollbar_button_color="#E0E0E0",  
            scrollbar_button_hover_color="#D0D0D0"
        ) 
        main_scroll.pack(fill="both", expand=True)

        header = ctk.CTkFrame(main_scroll, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(20, 10))
        ctk.CTkLabel(header, text="智能文档提取", font=("Microsoft YaHei", 24, "bold"), text_color="#333").pack(side="left")
        
        # ... (Step 1 & Step 2 UI 代码保持不变，节省篇幅，省略中间部分) ...
        # 请直接保留你原有的 render 代码，直到 Step 3 部分
        
        # ================= Step 1: 流程与输出 =================
        card_ai = ctk.CTkFrame(main_scroll, fg_color="white", corner_radius=10)
        card_ai.pack(fill="x", padx=20, pady=10)
        row_title_1 = ctk.CTkFrame(card_ai, fg_color="transparent")
        row_title_1.pack(fill="x", padx=20, pady=(15, 10))
        ctk.CTkLabel(row_title_1, text="Step 1. 流程控制", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(side="left")
        row1 = ctk.CTkFrame(card_ai, fg_color="transparent")
        row1.pack(fill="x", padx=20, pady=(0, 10))
        self.var_eye = ctk.BooleanVar(value=True)
        self.chk_eye = ctk.CTkCheckBox(row1, text="启用视觉引擎 (OCR)", variable=self.var_eye, text_color="#333")
        self.chk_eye.pack(side="left", padx=(0, 20))
        self.var_brain = ctk.BooleanVar(value=True)
        self.chk_brain = ctk.CTkCheckBox(row1, text="启用思考引擎 (提取)", variable=self.var_brain, text_color="#333")
        self.chk_brain.pack(side="left")
        row2 = ctk.CTkFrame(card_ai, fg_color="transparent")
        row2.pack(fill="x", padx=20, pady=(0, 15))
        self.var_save_raw = ctk.BooleanVar(value=False)
        self.chk_raw = ctk.CTkCheckBox(row2, text="保存 AI 原始回复 (调试)", variable=self.var_save_raw, text_color="#666", command=self.toggle_output_path)
        self.chk_raw.pack(side="left", padx=(0, 10))
        self.entry_out = ctk.CTkEntry(row2, placeholder_text="留痕文件保存路径...", width=250, fg_color="#FAFAFA", text_color="#333")
        self.btn_out = ctk.CTkButton(row2, text="📂", width=40, fg_color="#DDD", text_color="#333", hover_color="#CCC", command=self.select_output_dir)

        # ================= Step 2: 方案配置 =================
        card_schema = ctk.CTkFrame(main_scroll, fg_color="white", corner_radius=10)
        card_schema.pack(fill="x", padx=20, pady=10)
        row_title_2 = ctk.CTkFrame(card_schema, fg_color="transparent")
        row_title_2.pack(fill="x", padx=20, pady=(15, 10))
        ctk.CTkLabel(row_title_2, text="Step 2. 提取方案", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(side="left")
        row_sel = ctk.CTkFrame(card_schema, fg_color="transparent")
        row_sel.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(row_sel, text="当前方案:", text_color="#666").pack(side="left", padx=10)
        self.combo_schema = ctk.CTkComboBox(row_sel, values=["(自由模式 / 清空)"] + list(self.schemas.keys()) + ["+ 新建方案..."], command=self.on_schema_change, width=200, fg_color="white", border_color="#CCC", text_color="#333", button_color="#F0F0F0", button_hover_color="#E0E0E0", dropdown_fg_color="white", dropdown_text_color="#333")
        self.combo_schema.set("(自由模式 / 清空)")
        self.combo_schema.pack(side="left")
        ctk.CTkButton(row_sel, text="💾 保存", width=60, fg_color="#00b894", command=self.save_current_schema).pack(side="left", padx=5)
        ctk.CTkButton(row_sel, text="🗑️", width=40, fg_color="#ff7675", command=self.delete_current_schema).pack(side="left", padx=5)
        field_frame = ctk.CTkFrame(card_schema, fg_color="#FAFAFA", corner_radius=6)
        field_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(field_frame, text="结构化字段 (留空则启用自由模式)", text_color="#555", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        add_box = ctk.CTkFrame(field_frame, fg_color="transparent")
        add_box.pack(fill="x", padx=5, pady=5)
        self.entry_fname = ctk.CTkEntry(add_box, placeholder_text="字段名 (Key)", width=150, fg_color="white", text_color="black")
        self.entry_fname.pack(side="left")
        self.combo_ftype = ctk.CTkComboBox(add_box, values=["文本", "数字(小数)", "数字(整数)"], width=100, fg_color="white", text_color="black", button_color="#EEE")
        self.combo_ftype.set("文本")
        self.combo_ftype.pack(side="left", padx=5)
        ctk.CTkButton(add_box, text="+ 添加", width=60, command=self.add_field, fg_color="#007AFF").pack(side="left", padx=5)
        ctk.CTkButton(add_box, text="清空", width=60, command=self.clear_fields, fg_color="#999").pack(side="left", padx=5)
        self.scroll_fields = ctk.CTkScrollableFrame(field_frame, height=120, fg_color="white", scrollbar_button_color="#E0E0E0")
        self.scroll_fields.pack(fill="x", padx=10, pady=(0, 10))
        prompt_frame = ctk.CTkFrame(card_schema, fg_color="#FAFAFA", corner_radius=6)
        prompt_frame.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(prompt_frame, text="AI 任务指令 (Prompt)", text_color="#555", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        self.txt_prompt = ctk.CTkTextbox(prompt_frame, height=100, fg_color="white", text_color="#333", border_width=1, border_color="#DDD", font=("Consolas", 11))
        self.txt_prompt.insert("1.0", DEFAULT_USER_INSTRUCTION)
        self.txt_prompt.pack(fill="x", padx=10, pady=5)
        temp_box = ctk.CTkFrame(prompt_frame, fg_color="transparent")
        temp_box.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(temp_box, text="温度 (0.0严谨 - 1.0发散):", text_color="#666", font=("Arial", 11)).pack(side="left")
        self.slider_temp = ctk.CTkSlider(temp_box, from_=0.0, to=1.0, number_of_steps=10)
        self.slider_temp.set(0.1)
        self.slider_temp.pack(side="left", fill="x", expand=True, padx=10)
        self.lbl_temp_val = ctk.CTkLabel(temp_box, text="0.1", text_color="#007AFF", width=30)
        self.lbl_temp_val.pack(side="left")
        def update_temp_lbl(val): self.lbl_temp_val.configure(text=f"{val:.1f}")
        self.slider_temp.configure(command=update_temp_lbl)

        # ================= Step 3: 文件与执行 =================
        card_run = ctk.CTkFrame(main_scroll, fg_color="white", corner_radius=10)
        card_run.pack(fill="x", padx=20, pady=10)
        row_title_3 = ctk.CTkFrame(card_run, fg_color="transparent")
        row_title_3.pack(fill="x", padx=20, pady=(15, 10))
        ctk.CTkLabel(row_title_3, text="Step 3. 任务执行", font=("Microsoft YaHei", 14, "bold"), text_color="#333").pack(side="left")
        row_file = ctk.CTkFrame(card_run, fg_color="transparent")
        row_file.pack(fill="x", padx=10, pady=(0, 10))
        self.btn_file = ctk.CTkButton(row_file, text="📄 选择文件", command=self.select_file, width=120, fg_color="#007AFF")
        self.btn_file.pack(side="left", padx=10)
        self.lbl_file = ctk.CTkLabel(row_file, text="未选择", text_color="#666")
        self.lbl_file.pack(side="left", padx=10)
        
        btn_area = ctk.CTkFrame(main_scroll, fg_color="transparent")
        btn_area.pack(fill="x", padx=20, pady=20)
        
        self.btn_run = ctk.CTkButton(
            btn_area, text="🚀 执行工作流", command=self.run_process, height=50, font=("Microsoft YaHei", 18, "bold"), fg_color="#6c5ce7"
        )
        self.btn_run.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        # === 【修改点 2】停止按钮 command 指向本类新方法，新方法再调 app ===
        self.btn_stop = ctk.CTkButton(
            btn_area, text="🛑 终止任务", command=self.stop_process, height=50, font=("Microsoft YaHei", 18, "bold"), fg_color="#FF4757", state="disabled", width=140
        )
        self.btn_stop.pack(side="right")

        log_frame = ctk.CTkFrame(main_scroll, fg_color="white", corner_radius=10)
        log_frame.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(log_frame, text="系统日志", text_color="#333", font=("Consolas", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        self.log_box = ctk.CTkTextbox(log_frame, height=200, fg_color="white", text_color="black", font=("Consolas", 12), border_color="#DDD", border_width=1)
        self.log_box.pack(fill="x", padx=10, pady=(0, 10))

    # ... (UI Logic 保持不变) ...
    def toggle_output_path(self):
        if self.var_save_raw.get(): self.entry_out.pack(side="left", padx=5); self.btn_out.pack(side="left")
        else: self.entry_out.pack_forget(); self.btn_out.pack_forget()
    def select_output_dir(self):
        p = filedialog.askdirectory()
        if p: self.entry_out.delete(0, "end"); self.entry_out.insert(0, p)
    def on_schema_change(self, val):
        if val == "(自由模式 / 清空)":
            self.current_schema_fields = []; self.refresh_fields_ui()
            self.txt_prompt.delete("1.0", "end"); self.txt_prompt.insert("1.0", DEFAULT_USER_INSTRUCTION)
            self.slider_temp.set(0.1); self.lbl_temp_val.configure(text="0.1"); return
        if val == "+ 新建方案...":
            name = simpledialog.askstring("新建", "请输入方案名称:")
            if name:
                self.schemas[name] = {"name": name, "fields": [], "instruction": DEFAULT_USER_INSTRUCTION, "temperature": 0.1}
                self.combo_schema.configure(values=["(自由模式 / 清空)"] + list(self.schemas.keys()) + ["+ 新建方案..."])
                self.combo_schema.set(name); self.current_schema_fields = []; self.refresh_fields_ui()
                self.txt_prompt.delete("1.0", "end"); self.txt_prompt.insert("1.0", DEFAULT_USER_INSTRUCTION)
                self.slider_temp.set(0.1); self.lbl_temp_val.configure(text="0.1")
            else: self.combo_schema.set("(自由模式 / 清空)")
        elif val in self.schemas:
            data = self.schemas[val]; self.current_schema_fields = data.get("fields", []); self.refresh_fields_ui()
            self.txt_prompt.delete("1.0", "end"); self.txt_prompt.insert("1.0", data.get("instruction", DEFAULT_USER_INSTRUCTION))
            temp = data.get("temperature", 0.1); self.slider_temp.set(temp); self.lbl_temp_val.configure(text=f"{temp:.1f}")
    def refresh_fields_ui(self):
        for w in self.scroll_fields.winfo_children(): w.destroy()
        if not self.current_schema_fields: ctk.CTkLabel(self.scroll_fields, text="(当前为自由模式，AI 将直接回答)", text_color="#999").pack(pady=10); return
        for i, f in enumerate(self.current_schema_fields):
            row = ctk.CTkFrame(self.scroll_fields, fg_color="transparent", height=26); row.pack(fill="x", pady=2)
            left = ctk.CTkFrame(row, fg_color="transparent"); left.pack(side="left", padx=10)
            ctk.CTkLabel(left, text=f"{f['name']}", text_color="#333", anchor="w", width=120).pack(side="left")
            ctk.CTkLabel(left, text=f"[{f['type']}]", text_color="#666").pack(side="left", padx=5)
            ctk.CTkButton(row, text="×", width=20, height=20, fg_color="#FFF0F0", text_color="red", hover_color="#FFE0E0", command=lambda idx=i: self.delete_field(idx)).pack(side="right", padx=10)
    def add_field(self):
        name = self.entry_fname.get().strip()
        if not name: return
        self.current_schema_fields.append({"name": name, "type": self.combo_ftype.get(), "desc": name})
        self.refresh_fields_ui(); self.entry_fname.delete(0, "end")
    def delete_field(self, idx): self.current_schema_fields.pop(idx); self.refresh_fields_ui()
    def clear_fields(self): self.current_schema_fields = []; self.refresh_fields_ui()
    def save_current_schema(self):
        name = self.combo_schema.get()
        if name and name not in ["+ 新建方案...", "(自由模式 / 清空)"]:
            instruction = self.txt_prompt.get("1.0", "end-1c"); temp = self.slider_temp.get()
            LogicCore.save_schema(name, self.current_schema_fields, instruction, temp); messagebox.showinfo("成功", f"方案 [{name}] 已保存")
    def delete_current_schema(self):
        name = self.combo_schema.get()
        if name and name in self.schemas:
            if messagebox.askyesno("确认", f"删除方案 {name}?"):
                LogicCore.delete_schema(name); del self.schemas[name]; self.combo_schema.configure(values=["(自由模式 / 清空)"] + list(self.schemas.keys()) + ["+ 新建方案..."])
                self.combo_schema.set("(自由模式 / 清空)"); self.current_schema_fields = []; self.refresh_fields_ui()
    def select_file(self):
        p = filedialog.askopenfilename(filetypes=[("Documents", "*.pdf;*.png;*.jpg;*.jpeg")])
        if p: self.src_path = p; self.lbl_file.configure(text=os.path.basename(p))
    def log(self, msg): self.log_box.insert("end", f"> {msg}\n"); self.log_box.see("end")

    # === 【修改点 3】借力打力，让按钮触发 app 的停止逻辑 ===
    def stop_process(self):
        if hasattr(self, 'app'):
            # 直接调用主程序的停止方法，主程序会设置 stop_event
            self.app.stop_current_task()
        self.log("🛑 正在终止...")
        self.btn_stop.configure(state="disabled")

    # === 【修改点 4】启动任务 ===
    def run_process(self):
        if not self.src_path: messagebox.showwarning("提示", "请选择文件"); return
        
        use_eye = self.var_eye.get()
        use_brain = self.var_brain.get()
        instruction = self.txt_prompt.get("1.0", "end-1c")
        temperature = self.slider_temp.get()

        raw_output_dir = None
        if self.var_save_raw.get():
            raw_output_dir = self.entry_out.get().strip()
            if not raw_output_dir or not os.path.exists(raw_output_dir):
                messagebox.showwarning("提示", "无效的调试文件路径"); return

        # 申请全局红旗
        stop_event = None
        if hasattr(self, 'app'): 
            stop_event = self.app.register_task(self.module_index)
        
        # 兼容性处理：如果没有app注入（单独测试时），创建一个本地 event
        if not stop_event:
            stop_event = threading.Event()

        self.btn_run.configure(state="disabled", text="执行中...")
        self.btn_stop.configure(state="normal")
        self.log_box.delete("1.0", "end")

        def task():
            try:
                self.log(f"申请 AI 算力...")
                client_eye, model_eye = AIManager.get_client("vision")
                client_brain, model_brain = AIManager.get_client("brain")
                
                if use_eye and not client_eye: self.log("⚠️ 警告: 视觉引擎未配置，OCR 跳过")
                if use_brain and not client_brain: self.log("❌ 错误: 思考引擎未配置"); return

                if use_eye and client_eye: self.log(f"👁️ Eye: {model_eye}")
                if use_brain and client_brain: self.log(f"🧠 Brain: {model_brain} (T={temperature:.1f})")

                self.log("预处理文档...")
                inputs = []
                is_img_input = True
                if self.src_path.lower().endswith('.pdf'):
                    if use_eye: inputs = LogicCore.pdf_to_base64_images(self.src_path)
                    else:
                        inputs = LogicCore.pdf_to_text(self.src_path)
                        is_img_input = False
                else:
                    with open(self.src_path, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode("utf-8")
                        inputs = [f"data:image/jpeg;base64,{b64}"]

                # === 中断检测 1 ===
                if stop_event.is_set(): return

                self.log(f"开始处理 {len(inputs)} 页...")
                all_excel_data = [] 
                all_markdown_data = []
                
                with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                    futures = {
                        executor.submit(
                            process_pipeline_task, 
                            item, is_img_input, use_eye, use_brain, 
                            self.current_schema_fields,
                            instruction, temperature,
                            client_eye, model_eye,
                            client_brain, model_brain,
                            self.var_save_raw.get() 
                        ): i for i, item in enumerate(inputs)
                    }
                    
                    for future in concurrent.futures.as_completed(futures):
                        # === 中断检测 2 ===
                        if stop_event.is_set():
                            executor.shutdown(wait=False, cancel_futures=True)
                            self.log("🛑 任务已强制终止"); break

                        idx = futures[future]
                        structured_rows, raw_response, err, usage = future.result()
                        
                        if raw_output_dir:
                            ts = int(time.time())
                            if raw_response:
                                with open(os.path.join(raw_output_dir, f"P{idx+1}_Response_{ts}.md"), "w", encoding="utf-8") as f: f.write(raw_response)

                        if err:
                            self.log(f"[×] P{idx+1}: {err}")
                        else:
                            if raw_response:
                                all_markdown_data.append(f"## Page {idx+1}\n\n{raw_response}\n\n---")
                            if structured_rows:
                                count = len(structured_rows)
                                self.log(f"[√] P{idx+1} 结构化提取: {count} 条")
                                all_excel_data.extend(structured_rows)
                            else:
                                self.log(f"[√] P{idx+1} 分析完成 (无结构化数据)")

                # === 中断检测 3 ===
                if stop_event.is_set(): return

                if not all_excel_data and not all_markdown_data:
                    self.log("无有效输出。")
                    return

                # === 导出 (由于要弹窗选择目录，建议放到主线程，但这里简单处理直接弹) ===
                # 注意：如果线程里弹 filedialog 卡死，可改用 self.app.after 调度
                output_dir = filedialog.askdirectory(title="选择结果保存目录")
                if output_dir:
                    ts = int(time.time())
                    md_path = os.path.join(output_dir, f"AI_Analysis_{ts}.md")
                    with open(md_path, "w", encoding="utf-8") as f:
                        f.write("\n".join(all_markdown_data))
                    self.log(f"文档分析已保存: {os.path.basename(md_path)}")
                    
                    if all_excel_data:
                        xls_path = os.path.join(output_dir, f"Extracted_Data_{ts}.xlsx")
                        pd.DataFrame(all_excel_data).to_excel(xls_path, index=False)
                        self.log(f"结构化数据已保存: {os.path.basename(xls_path)}")
                        os.startfile(output_dir)
                    
                    messagebox.showinfo("完成", "任务执行完毕")

            except Exception as e:
                self.log(f"Error: {e}")
            finally:
                # 销假 & 恢复按钮
                if hasattr(self, 'app'): self.app.finish_task(self.module_index)
                self.btn_run.configure(state="normal", text="🚀 执行工作流")
                self.btn_stop.configure(state="disabled")

        threading.Thread(target=task, daemon=True).start()