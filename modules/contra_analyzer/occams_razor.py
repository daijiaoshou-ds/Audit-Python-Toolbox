class OccamsRazor:
    """
    奥卡姆剃刀剪枝器 v3.0
    公式：Score = 100 - (行数 * 1) - (分拆惩罚)
    """
    
    @staticmethod
    def score_solution(solution):
        """
        solution: {借方: {贷方: 金额}}
        """
        score = 100
        
        # 1. 统计借方数量(n) 和 贷方数量(m)
        d_split_counts = {} 
        c_split_counts = {}
        
        all_d = set(solution.keys())
        all_c = set()
        
        # 计算总行数 (连接数)
        total_lines = 0
        
        for d, c_map in solution.items():
            valid_c_links = [c for c, amt in c_map.items() if abs(amt) > 0.001]
            d_split_counts[d] = len(valid_c_links)
            total_lines += len(valid_c_links)
            
            for c in valid_c_links:
                all_c.add(c)
                c_split_counts[c] = c_split_counts.get(c, 0) + 1
                
        n = len(all_d)
        m = len(all_c)
        
        # === 扣分项 1: 行数惩罚 ===
        score -= total_lines * 1
        
        # === 扣分项 2: 分拆惩罚 ===
        # 规则：数量多的一方做 Driver (遍历方)，数量少的是 Bucket
        # Driver 被分拆 -> 扣 5 分
        # Bucket 被分拆 -> 扣 1 分
        
        debit_is_driver = (n >= m)
        
        penalty_driver = 5
        penalty_bucket = 1
        
        if debit_is_driver:
            # 借方是 Driver
            for count in d_split_counts.values():
                if count > 1: score -= (count - 1) * penalty_driver
            for count in c_split_counts.values():
                if count > 1: score -= (count - 1) * penalty_bucket
        else:
            # 贷方是 Driver
            for count in c_split_counts.values():
                if count > 1: score -= (count - 1) * penalty_driver
            for count in d_split_counts.values():
                if count > 1: score -= (count - 1) * penalty_bucket
        
        return round(score, 2)

    @staticmethod
    def rank_solutions(solutions):
        """仅按奥卡姆得分排序 (辅助用，主逻辑在 UI)"""
        if not solutions: return [], []
        scored_items = []
        for sol in solutions:
            score = OccamsRazor.score_solution(sol)
            scored_items.append((sol, score))
        scored_items.sort(key=lambda x: x[1], reverse=True)
        return [x[0] for x in scored_items], [x[1] for x in scored_items]