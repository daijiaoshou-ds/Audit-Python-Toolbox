from .ui import ContraAnalyzerUI

# 为了适配 main.py 的加载逻辑，我们需要把 UI 类包装成符合要求的 Module 类
# 你的 main.py 期望的是: module.name, module.render(parent), module.app (注入)

class ContraAnalyzerModule:
    def __init__(self):
        self.ui = ContraAnalyzerUI()
        self.name = self.ui.name
        # self.kb = KnowledgeBase() <--- 这行删掉！UI 里已经有了，这里不需要。
        
    def render(self, parent_frame):
        # 传递 app 实例给内部 UI，以便能获取 stop_event
        if hasattr(self, 'app'):
            self.ui.app = self.app
            self.ui.module_index = self.module_index
            
        self.ui.render(parent_frame)