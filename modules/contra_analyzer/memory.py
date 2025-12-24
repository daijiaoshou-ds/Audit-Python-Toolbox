import os
import json
from modules.path_manager import get_user_data_dir
from .occams_razor import OccamsRazor

class KnowledgeBase:
    """
    记忆库 v9.1 (指纹排序修复版)
    核心修复：指纹生成时，先提取所有连接对，清洗后再统一排序。
    确保 [Solver生成的带后缀数据] 和 [Excel导入的不带后缀数据] 能生成完全一致的指纹。
    """
    def __init__(self):
        self.file_path = os.path.join(get_user_data_dir(), "contra_memory_ema.json")
        self.learning_rate = 0.6  
        self.beta_factor = 0.5    
        self.memory = self._load()

    def _load(self):
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except: pass
        return {}

    def save(self):
        with open(self.file_path, 'w', encoding='utf-8') as f:
            json.dump(self.memory, f, ensure_ascii=False, indent=2)

    def clear_memory(self):
        self.memory = {}
        if os.path.exists(self.file_path):
            os.remove(self.file_path)

    def _generate_fingerprint(self, solution):
        """
        生成【规范化结构指纹】。
        1. 遍历 solution，提取所有非零连接 (d, c)。
        2. 清洗 d 和 c (去除后缀)。
        3. 将所有 (clean_d, clean_c) 放入列表。
        4. 对列表进行【统一排序】。
        5. 拼接字符串。
        """
        connections = []
        
        for d, c_map in solution.items():
            for c, amt in c_map.items():
                if abs(amt) > 0.001:
                    # 清洗 Key
                    clean_d = str(d).split('__')[0].strip()
                    clean_c = str(c).split('__')[0].strip()
                    
                    # 过滤无效数据
                    if not clean_d or not clean_c: continue
                    if clean_d.lower() == 'nan' or clean_c.lower() == 'nan': continue
                    
                    connections.append(f"{clean_d}->{clean_c}")
        
        # === 核心：统一排序 ===
        # 无论输入字典的 Key 顺序如何，这里强制按字符串内容排序
        connections.sort()
        
        return "|".join(connections)

    def get_memory_score(self, pattern_name, solution):
        fp = self._generate_fingerprint(solution)
        if pattern_name in self.memory:
            return self.memory[pattern_name].get(fp, 0.5)
        return 0.5

    def update_memory(self, pattern_name, all_solutions, selected_solution):
        """
        EMA 更新
        """
        if pattern_name not in self.memory:
            self.memory[pattern_name] = {}
        
        # 生成标准指纹
        target_fp = self._generate_fingerprint(selected_solution)
        
        # 记录本次出现的所有指纹
        seen_fps = set()
        
        for sol in all_solutions:
            fp = self._generate_fingerprint(sol)
            if not fp: continue 
            seen_fps.add(fp)
            
            m_old = self.memory[pattern_name].get(fp, 0.5)
            
            # 命中
            if fp == target_fp:
                reward = 1.0 
            else:
                reward = 0.0
            
            m_new = m_old * (1 - self.learning_rate) + reward * self.learning_rate
            self.memory[pattern_name][fp] = round(m_new, 4)
            
        self.save()

    def update_memory_by_fingerprint(self, pattern_name, all_solutions, target_fingerprint):
        """
        UI 直接传入已生成好的 target_fingerprint
        """
        if pattern_name not in self.memory:
            self.memory[pattern_name] = {}
            
        for sol in all_solutions:
            fp = self._generate_fingerprint(sol)
            if not fp: continue
            
            m_old = self.memory[pattern_name].get(fp, 0.5)
            
            reward = 1.0 if fp == target_fingerprint else 0.0
            
            m_new = m_old * (1 - self.learning_rate) + reward * self.learning_rate
            self.memory[pattern_name][fp] = round(m_new, 4)
            
        self.save()

    def calculate_total_score(self, razor_score, memory_score):
        return round(razor_score * (1 + self.beta_factor * memory_score), 2)

    def rank_solutions(self, solutions, pattern_name=""):
        if not solutions: return []
        
        scored = []
        for sol in solutions:
            r = OccamsRazor.score_solution(sol)
            m = self.get_memory_score(pattern_name, sol) if pattern_name else 0.5
            total = self.calculate_total_score(r, m)
            scored.append((total, sol))
            
        scored.sort(key=lambda x: x[0], reverse=True)
        return [x[1] for x in scored]