import itertools
from collections import defaultdict

class SolutionOptimizer:
    """
    规划求解优化器
    核心思想：变异性分析 + 奥卡姆剃刀(最简解优先)
    """
    
    @staticmethod
    def get_recommended_subset(solutions, debits_dict, credits_dict):
        """
        在所有方案中，找出符合"最简数学逻辑"的方案子集。
        :return: 推荐的方案列表 (subset of solutions)
        """
        if not solutions or len(solutions) <= 1:
            return solutions

        # 复制一份作为候选池
        candidates = solutions[:] 
        
        # 1. 计算每个借方科目的"变异性"
        variance_map = {} 
        all_d_keys = list(debits_dict.keys())
        
        for d_key in all_d_keys:
            seen_splits = set()
            for sol in candidates:
                if d_key in sol:
                    # 转 tuple 用于哈希去重
                    split = sol[d_key]
                    valid_split = tuple(sorted(
                        [(k, v) for k, v in split.items() if abs(v) > 0.001],
                        key=lambda x: x[0]
                    ))
                    seen_splits.add(valid_split)
            variance_map[d_key] = len(seen_splits)

        # 2. 按变异性从高到低排序
        sorted_keys = sorted(
            [k for k, v in variance_map.items() if v > 1],
            key=lambda k: variance_map[k],
            reverse=True
        )
        
        if not sorted_keys:
            return candidates # 没差异，推荐全部

        # 3. 逐个攻破，进行筛选
        for d_key in sorted_keys:
            if len(candidates) <= 1: break 
            
            target_amt = debits_dict[d_key]
            
            # === 数学暴力规划求解 (寻找最简组合) ===
            best_combination = SolutionOptimizer._find_simplest_combination(target_amt, credits_dict)
            
            if not best_combination:
                continue 
            
            # 4. 筛选：只保留符合这个"最简组合"的方案
            filtered = []
            for sol in candidates:
                actual_split = sol[d_key]
                actual_targets = [k for k, v in actual_split.items() if abs(v) > 0.001]
                
                # 判定：拆分对象集合是否一致
                if set(actual_targets) == set(best_combination):
                    filtered.append(sol)
            
            # 只有当筛选结果不为空时，才缩小候选池
            # (防止数学求解过于理想化，实际方案里没有)
            if filtered:
                candidates = filtered

        return candidates

    @staticmethod
    def _find_simplest_combination(target, pool_dict):
        """
        寻找和为 target 的最简组合 (1项 > 2项 > 3项)
        """
        items = [(k, v) for k, v in pool_dict.items() if abs(v) > 0.001]
        n = len(items)
        
        # 尝试 k = 1 到 3
        for r in range(1, min(n + 1, 4)):
            for comb in itertools.combinations(items, r):
                s = sum(x[1] for x in comb)
                if abs(s - target) < 0.001:
                    return [x[0] for x in comb]
        return None