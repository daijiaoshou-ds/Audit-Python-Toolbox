class OccamsRazor:
    """
    奥卡姆剃刀剪枝器
    核心思想：如无必要，勿增实体。
    在会计分录中，这意味着：借贷关系的连接数越少，该方案越可能是真实的业务逻辑。
    """
    
    @staticmethod
    def score_solution(solution):
        """
        计算方案的复杂度得分。
        solution 结构: {借方科目: {贷方科目: 金额}}
        
        计分规则：
        1. 初始分：100
        2. 连接惩罚：每存在一个非零金额的 (借->贷) 关系，扣 1 分。
        3. 碎片惩罚（可选）：如果金额是很碎的小数，额外扣分（暂时不加，保持简单）。
        """
        score = 100
        connection_count = 0
        
        for d_subj, c_map in solution.items():
            for c_subj, amt in c_map.items():
                # 只有真实存在的连接才算
                if abs(amt) > 0.001:
                    connection_count += 1
                    score -= 1 # 每多一条线，扣1分
        
        # 也可以返回负的连接数，效果一样。这里返回标准化分数。
        return score, connection_count

    @staticmethod
    def rank_solutions(solutions):
        """
        对方案列表进行排序，最简单的排前面。
        返回: (排序后的方案列表, 对应的分数列表)
        """
        if not solutions:
            return [], []
            
        # 列表推导式计算所有得分
        # item: (solution, score, connection_count)
        scored_items = []
        for sol in solutions:
            score, count = OccamsRazor.score_solution(sol)
            scored_items.append((sol, score))
            
        # 排序：分数从大到小 (即连接数从小到大)
        # 如果分数相同，保持原有顺序 (稳定排序)
        scored_items.sort(key=lambda x: x[1], reverse=True)
        
        sorted_solutions = [item[0] for item in scored_items]
        sorted_scores = [item[1] for item in scored_items]
        
        return sorted_solutions, sorted_scores