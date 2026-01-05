import PyInstaller.__main__
import os
import shutil
import time

def build():
    exe_name = "Python工具箱"
    
    # 获取路径
    project_root = os.path.abspath(".")
    dist_dir = os.path.join(project_root, "dist")
    build_dir = os.path.join(project_root, "build")
    spec_file = os.path.join(project_root, f"{exe_name}.spec")
    
    # 目标 EXE 在 dist 里的路径
    src_exe = os.path.join(dist_dir, f"{exe_name}.exe")
    # 最终 EXE 要放的根目录路径
    dst_exe = os.path.join(project_root, f"{exe_name}.exe")

    # 1. 清理旧构建 & 旧 EXE
    print("🧹 正在清理旧文件...")
    if os.path.exists(build_dir): shutil.rmtree(build_dir)
    if os.path.exists(dist_dir): shutil.rmtree(dist_dir)
    if os.path.exists(spec_file): os.remove(spec_file)
    if os.path.exists("ExcelToolsPro.spec"): os.remove("ExcelToolsPro.spec")
    
    # 如果根目录下已经有一个旧的 EXE，先删掉，防止覆盖报错
    if os.path.exists(dst_exe):
        try:
            os.remove(dst_exe)
            print(f"   已删除根目录下的旧版本: {exe_name}.exe")
        except Exception as e:
            print(f"❌ 无法删除旧 EXE (可能正在运行?): {e}")
            return

    print(f"🚀 开始打包: {exe_name} ...")
    print("⏳ 请耐心等待，NLP 依赖较多...")

    params = [
        'main.py',
        f'--name={exe_name}',
        '--onefile',
        '--noconsole',
        '--clean',
        '--icon=assets/icon.ico',
        '--noupx',
        
        # 资源文件 (只打包小的)
         '--add-data=assets/fonts/simsun.ttc;assets/fonts', 
        '--add-data=assets/icon.ico;assets',    
        
        # 依赖收集
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
        # === 【新增】Scipy 相关 (匈牙利算法必须) ===
        '--collect-all=scipy', 
        '--collect-all=polars',      # 【新增】Polars引擎
        '--collect-all=pyarrow',     # 【新增】Polars转换依赖
        # --- 隐式导入 ---
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=pydantic.deprecated.decorator',
        '--hidden-import=sklearn.utils._typedefs',
        '--hidden-import=sklearn.neighbors._partition_nodes',
        '--hidden-import=sklearn.tree',
        '--hidden-import=sklearn.ensemble',
        '--hidden-import=sentence_transformers',
        '--hidden-import=huggingface_hub',
        # === 【新增】Scipy 隐式导入 ===
        '--hidden-import=scipy.special.cython_special',
        '--hidden-import=scipy.spatial.transform._rotation_groups',
    ]

    try:
        PyInstaller.__main__.run(params)
    except Exception as e:
        print(f"\n❌ 打包失败: {e}")
        return

    print("\n📦 打包完成，正在执行自动化部署...")

    # === 【核心修改】自动搬运 EXE 到根目录 ===
    if os.path.exists(src_exe):
        # 1. 移动文件
        shutil.move(src_exe, dst_exe)
        print(f"✅ 成功！已将 EXE 移动到项目根目录: \n   -> {dst_exe}")
        
        # 2. 清理 dist 和 build 文件夹 (强迫症福音)
        time.sleep(1) # 等待文件句柄释放
        if os.path.exists(dist_dir): shutil.rmtree(dist_dir)
        if os.path.exists(build_dir): shutil.rmtree(build_dir)
        if os.path.exists(spec_file): os.remove(spec_file)
        print("🧹 已清理临时构建文件 (dist/build/spec)")
        
        print("\n" + "="*50)
        print("🎉 全部搞定！")
        print("现在直接在根目录下双击【Python工具箱.exe】即可运行。")
        print("它会自动加载旁边的 assets 文件夹和 user_data 配置。")
        print("="*50)
    else:
        print("❌ 错误：在 dist 中未找到生成的 EXE，打包可能未成功。")

if __name__ == "__main__":
    build()