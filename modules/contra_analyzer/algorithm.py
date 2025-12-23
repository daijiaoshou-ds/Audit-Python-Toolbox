import itertools
import hashlib
import time
from collections import defaultdict

class ExhaustiveSolver:
    """
    穷举算法 v6.1 (无限制数学版 + 精度增强)
    修复：允许负数 Driver 分配给正数 Bucket (反向增加容量)，解决正负混杂无解问题。
    """
    
    @staticmethod
    def calculate_combinations(debit_ledger, credit_ledger, max_solutions=200, timeout=5.0):
        start_time = time.time()
        
        # 1. 预处理
        # 严格保留原始符号，不过滤负数
        debits = {k: v for k, v in debit_ledger.items() if abs(v) > 0.001}
        credits = {k: v for k, v in credit_ledger.items() if abs(v) > 0.001}
        
        # 2. 智能转置 (数量优先)
        n_debits = len(debits)
        n_credits = len(credits)
        is_transposed = False
        
        if n_credits > n_debits:
            is_transposed = True
            drivers = credits
            buckets = debits
        else:
            drivers = debits
            buckets = credits

        # 3. 核心计算
        raw_results, is_timeout = ExhaustiveSolver._core_solve(drivers, buckets, max_solutions, timeout, start_time)

        # 4. 结果还原
        deduped_results = []
        seen_hashes = set()

        for result in raw_results:
            processed_res = {}
            for d_key in debits.keys(): processed_res[d_key] = {}

            if is_transposed:
                for c_key, d_map in result.items():
                    for d_key, amount in d_map.items():
                        if abs(amount) > 0.001:
                            processed_res[d_key][c_key] = amount
            else:
                for d_key, c_map in result.items():
                    for c_key, amount in c_map.items():
                        if abs(amount) > 0.001:
                            processed_res[d_key][c_key] = amount

            # 签名去重
            temp_sig = []
            sorted_d = sorted(processed_res.keys())
            for d in sorted_d:
                sorted_c = sorted(processed_res[d].keys())
                for c in sorted_c:
                    amt = processed_res[d][c]
                    if abs(amt) > 0.001:
                        temp_sig.append(f"{d}:{c}:{round(amt, 4)}")
            
            sig_str = "|".join(temp_sig)
            res_hash = hashlib.md5(sig_str.encode()).hexdigest()
            
            if res_hash not in seen_hashes:
                seen_hashes.add(res_hash)
                deduped_results.append(processed_res)

        return deduped_results, is_timeout

    @staticmethod
    def _core_solve(drivers_dict, buckets_dict, max_sol, timeout, start_time):
        # === 排序策略：按【代数大小】从小到大 ===
        # 负数 (-521.5) 会排在 正数 (10670) 前面
        # 这样 -521.5 先去填坑 (或挖坑)，符合你的推导
        driver_items = sorted(drivers_dict.items(), key=lambda x: x[1], reverse=False)
        bucket_items = sorted(list(buckets_dict.items()), key=lambda x: x[1], reverse=False)
        
        results = []
        is_timeout = [False]

        def generate_combinations(target_amt, available_buckets):
            valid_splits = []
            n = len(available_buckets)
            seen = set()
            
            # A. 全匹配 (Subset Sum)
            # 寻找 buckets 的子集，其和等于 target_amt
            for r in range(1, n + 1):
                for indices in itertools.combinations(range(n), r):
                    # 精度修正
                    subset_sum = sum(available_buckets[i][1] for i in indices)
                    subset_sum = round(subset_sum, 4)
                    
                    if abs(subset_sum - round(target_amt, 4)) < 0.001:
                        split_map = {}
                        for i in indices:
                            split_map[available_buckets[i][0]] = available_buckets[i][1]
                        
                        key = tuple(sorted((k, round(v, 4)) for k,v in split_map.items()))
                        if key not in seen:
                            seen.add(key)
                            valid_splits.append(split_map)

            # B. 部分匹配 (至多1个非边界)
            for i in range(n):
                partial_name, partial_cap = available_buckets[i]
                
                # 剩余需要的钱
                others_indices = [x for x in range(n) if x != i]
                n_others = len(others_indices)
                
                for r in range(n_others + 1):
                    for sub_indices in itertools.combinations(others_indices, r):
                        current_sum = sum(available_buckets[k][1] for k in sub_indices)
                        current_sum = round(current_sum, 4)
                        
                        needed = target_amt - current_sum
                        needed = round(needed, 4)
                        
                        # === 核心修正 ===
                        # 只要 needed 不为 0，且不等于 partial_cap (那是全匹配)，就允许！
                        # 不再检查 needed 是否在 (0, partial_cap) 之间。
                        # 因为负数分配给正数 Bucket 是合法的 (相当于扩容)。
                        
                        if abs(needed) > 0.001 and abs(needed - partial_cap) > 0.001:
                            split_map = {}
                            for k in sub_indices:
                                split_map[available_buckets[k][0]] = available_buckets[k][1]
                            split_map[partial_name] = needed
                            
                            key = tuple(sorted((k, round(v, 4)) for k,v in split_map.items()))
                            if key not in seen:
                                seen.add(key)
                                valid_splits.append(split_map)
            return valid_splits

        def dfs(d_idx, current_allocations, current_buckets):
            if len(results) >= max_sol: return
            if time.time() - start_time > timeout:
                is_timeout[0] = True; return

            if d_idx == len(driver_items):
                # 检查 Buckets 是否清零 (精度修正)
                remain = sum(amt for _, amt in current_buckets)
                if abs(remain) < 0.01: # 放宽一点点总误差容忍度
                    results.append(current_allocations)
                return

            driver_name, driver_amt = driver_items[d_idx]
            
            possible_splits = generate_combinations(driver_amt, current_buckets)
            
            # 如果没找到方案，剪枝
            if not possible_splits: return

            for split in possible_splits:
                if len(results) >= max_sol: return
                
                next_alloc = current_allocations.copy()
                next_alloc[driver_name] = split
                
                # 更新剩余 Buckets
                next_buckets = []
                for b_name, b_amt in current_buckets:
                    used = split.get(b_name, 0)
                    remain = b_amt - used
                    remain = round(remain, 4) # 保持精度
                    
                    # 只要绝对值还有剩余，就带入下一轮
                    # 即使是负数 Driver 把正数 Bucket 变成了更大的正数，也要带下去
                    if abs(remain) > 0.001:
                        next_buckets.append((b_name, remain))
                
                dfs(d_idx + 1, next_alloc, next_buckets)

        dfs(0, {}, bucket_items)
        return results, is_timeout[0]