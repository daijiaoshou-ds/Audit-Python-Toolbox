import os
import json
from collections import defaultdict
from modules.path_manager import get_user_data_dir

class KnowledgeBase:
    """
    记忆库 v3.0 (加权矩阵版)
    逻辑：记录 [借方科目] 与 [贷方科目] 的亲密度分数。
    机制：
    1. 累加制：分数越高，代表关系越稳固。
    2. 暴击制：用户手动导入的规则，给予高权重(例如+500)，实现快速纠错/覆盖。
    """
    def __init__(self):
        self.file_path = os.path.join(get_user_data_dir(), "contra_matrix_v3.json")
        self.matrix = self._load() # 结构: {借方: {贷方: 分数}}

    def _load(self):
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {}

    def save(self):
        with open(self.file_path, 'w', encoding='utf-8') as f:
            json.dump(self.matrix, f, ensure_ascii=False, indent=2)

    def learn_from_solution(self, solution, weight=500):
        """
        学习一个方案
        weight: 权重。默认为 500 (暴击)，确保用户指定的方案下次一定排第一。
        """
        for d_subj, c_map in solution.items():
            # 清洗科目名 (去掉 _idx_D 后缀)
            clean_d = d_subj.split('__')[0]
            
            for c_subj, amt in c_map.items():
                if abs(amt) > 0.001:
                    clean_c = c_subj.split('__')[0]
                    
                    if clean_d not in self.matrix:
                        self.matrix[clean_d] = {}
                    
                    # 累加分数
                    current_score = self.matrix[clean_d].get(clean_c, 0)
                    self.matrix[clean_d][clean_c] = current_score + weight
        
        self.save()

    def rank_solutions(self, solutions):
        """
        给方案列表排序
        计算每个方案的"总亲密度"，分数高的排前面
        """
        if not solutions: return []
        
        scored_solutions = []
        for sol in solutions:
            score = 0
            for d_subj, c_map in sol.items():
                clean_d = d_subj.split('__')[0]
                
                if clean_d in self.matrix:
                    for c_subj, amt in c_map.items():
                        if abs(amt) > 0.001:
                            clean_c = c_subj.split('__')[0]
                            # 查表加分
                            relation_score = self.matrix[clean_d].get(clean_c, 0)
                            score += relation_score
            
            scored_solutions.append((score, sol))
        
        # 降序排列 (分数高的在通过)
        scored_solutions.sort(key=lambda x: x[0], reverse=True)
        
        # 返回排序后的方案
        return [item[1] for item in scored_solutions]