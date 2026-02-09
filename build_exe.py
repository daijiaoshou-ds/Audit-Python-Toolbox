import PyInstaller.__main__
import os
import shutil
import time

def build():
    exe_name = "基米工具箱"
    project_root = os.path.abspath(".")
    
    # 定义输出目录
    dist_dir = os.path.join(project_root, "dist")
    build_dir = os.path.join(project_root, "build")
    spec_file = os.path.join(project_root, f"{exe_name}.spec")
    
    # 最终文件夹路径
    output_folder = os.path.join(dist_dir, exe_name) 

    print("🧹 清理旧构建...")
    if os.path.exists(build_dir): shutil.rmtree(build_dir)
    if os.path.exists(dist_dir): shutil.rmtree(dist_dir)
    if os.path.exists(spec_file): os.remove(spec_file)
    if os.path.exists("ExcelToolsPro.spec"): os.remove("ExcelToolsPro.spec")

    print(f"🚀 开始打包 (文件夹模式): {exe_name} ...")

    params = [
        'main.py',
        f'--name={exe_name}',
        '--onedir',            # <--- 【核心修改】由 --onefile 改为 --onedir
        '--noconsole',
        '--clean',
        '--icon=assets/icon.ico',
        '--noupx',
        
        # 资源文件 (modules 必须打包进去)
        '--add-data=modules;modules', 
        
        # ... (以下所有依赖收集 collect-all 和 hidden-import 保持完全不变) ...
        '--collect-all=customtkinter',
        '--collect-all=rembg',
        '--collect-all=onnxruntime',
        '--collect-all=pandas',
        '--collect-all=openpyxl',
        '--collect-all=pymupdf',
        '--collect-all=sklearn',
        '--collect-all=torch',
        '--collect-all=sentence_transformers',
        '--collect-all=jieba',
        '--collect-all=scipy', 
        '--collect-all=polars',
        '--collect-all=pyarrow',

        # 【新增】强制收集依赖库的 Metadata (修复 Transformers 版本检查报错)
        # ==================================================
        '--copy-metadata=tqdm',
        '--copy-metadata=regex',
        '--copy-metadata=requests',
        '--copy-metadata=tokenizers',
        '--copy-metadata=filelock',
        '--copy-metadata=huggingface_hub',
        '--copy-metadata=safetensors',
        '--copy-metadata=transformers',
        '--copy-metadata=sentence_transformers',
        '--copy-metadata=numpy',
        # ==================================================
        
        '--hidden-import=numpy',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=csv',
        '--hidden-import=json',
        '--hidden-import=difflib',
        '--hidden-import=re',
        '--hidden-import=hashlib',
        '--hidden-import=uuid',
        '--hidden-import=openai',
        '--hidden-import=tiktoken',
        '--hidden-import=torch',
        '--hidden-import=sklearn',
        '--hidden-import=sentence_transformers',
        '--hidden-import=huggingface_hub',
        '--hidden-import=PIL',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=rembg',
        '--hidden-import=onnxruntime',
        '--hidden-import=fitz',
        '--hidden-import=pdfplumber',
        '--hidden-import=reportlab',
        '--hidden-import=docx',
        '--hidden-import=docx2pdf',
        '--hidden-import=python_calamine',  
        '--hidden-import=xlrd',
        '--hidden-import=xlwt',
        '--hidden-import=xlsxwriter',
        '--hidden-import=fastexcel',
        '--hidden-import=threading',
        '--hidden-import=concurrent.futures',
        '--hidden-import=win32com',
        '--hidden-import=win32com.client',
        '--hidden-import=pythoncom',
        '--hidden-import=pydantic',
        '--hidden-import=pydantic.deprecated.decorator',
        '--hidden-import=sklearn.utils._typedefs',
        '--hidden-import=sklearn.neighbors._partition_nodes',
        '--hidden-import=sklearn.tree',
        '--hidden-import=sklearn.ensemble',
        '--hidden-import=scipy.special.cython_special',
        '--hidden-import=scipy.spatial.transform._rotation_groups',
        '--hidden-import=scipy.optimize',
    ]

    PyInstaller.__main__.run(params)

    print("\n📦 打包完成，正在组装最终文件夹...")

    # === 组装最终交付文件夹 ===
    # 目标结构:
    # dist/基米工具箱/
    #   ├── 基米工具箱.exe
    #   ├── assets/  <-- 我们手动拷进去
    #   └── _internal/ (依赖库)

    if os.path.exists(output_folder):
        # 1. 复制 assets 文件夹进去
        print("   正在复制资源文件 (assets)...")
        dest_assets = os.path.join(output_folder, "assets")
        if os.path.exists("assets"):
            # 如果目标存在先删再拷
            if os.path.exists(dest_assets): shutil.rmtree(dest_assets)
            shutil.copytree("assets", dest_assets)
        
        # 2. (可选) 复制 user_data 进去
        # dest_user = os.path.join(output_folder, "user_data")
        # if not os.path.exists(dest_user): os.makedirs(dest_user)

        print("\n" + "="*50)
        print("🎉 全部搞定！")
        print(f"请查看文件夹: {output_folder}")
        print("直接把这个【基米工具箱】文件夹发给同事即可！")
        print("="*50)
        
        # 自动打开文件夹
        os.startfile(dist_dir)

if __name__ == "__main__":
    build()