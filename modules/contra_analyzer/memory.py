import os
import json
from collections import defaultdict
from modules.path_manager import get_user_data_dir

class KnowledgeBase:
    def __init__(self):
        self.file_path = os.path.join(get_user_data_dir(), "contra_matrix.json")
        self.matrix = self._load()

    def _load(self):
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {} # 结构: { "借方科目": { "贷方科目": score } }

    def save(self):
        with open(self.file_path, 'w', encoding='utf-8') as f:
            json.dump(self.matrix, f, ensure_ascii=False, indent=2)

    def learn_from_solution(self, solution):
        """
        从用户选择的方案中学习
        solution: {借方: {贷方: 金额}}
        """
        for d_subj, c_map in solution.items():
            # 这里的 d_subj 可能带 _idx 后缀，清洗一下
            clean_d = d_subj.split('_')[0]
            
            for c_subj, amt in c_map.items():
                if abs(amt) > 0.001:
                    clean_c = c_subj.split('_')[0]
                    
                    # 增加权重
                    if clean_d not in self.matrix:
                        self.matrix[clean_d] = {}
                    
                    # 简单加分策略：每次 +1
                    self.matrix[clean_d][clean_c] = self.matrix[clean_d].get(clean_c, 0) + 1
        
        self.save()

    def rank_solutions(self, solutions):
        """
        给方案列表排序，得分高的排前面
        """
        if not solutions: return []
        
        scored_solutions = []
        for sol in solutions:
            score = 0
            for d_subj, c_map in sol.items():
                clean_d = d_subj.split('_')[0]
                if clean_d in self.matrix:
                    for c_subj, amt in c_map.items():
                        if abs(amt) > 0.001:
                            clean_c = c_subj.split('_')[0]
                            # 累加亲密度分数
                            score += self.matrix[clean_d].get(clean_c, 0)
            
            scored_solutions.append((score, sol))
        
        # 按分数降序排列
        scored_solutions.sort(key=lambda x: x[0], reverse=True)
        
        # 返回排序后的 solutions
        return [item[1] for item in scored_solutions]