import itertools
import hashlib
import time
from collections import defaultdict

class ExhaustiveSolver:
    """
    穷举算法 v3.2 (同名科目符号冲突修复版)
    """
    
    @staticmethod
    def calculate_combinations(debit_ledger, credit_ledger, max_solutions=200, timeout=2.0):
        start_time = time.time()
        
        # --- 1. 预处理 (拆分符号映射) ---
        debit_signs = {}   # 专门存借方科目的符号
        credit_signs = {}  # 专门存贷方科目的符号
        
        # 借方处理
        debits = {}
        for account, amount in debit_ledger.items():
            debit_signs[account] = 1 if amount >= 0 else -1
            if abs(amount) > 0.001:
                debits[account] = abs(amount)

        # 贷方处理
        credits = {}
        for account, amount in credit_ledger.items():
            credit_signs[account] = 1 if amount >= 0 else -1
            if abs(amount) > 0.001:
                credits[account] = abs(amount)
        
        # --- 2. 智能转置 ---
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

        # --- 3. 核心计算 ---
        raw_results, is_timeout = ExhaustiveSolver._core_solve(drivers, buckets, max_solutions, timeout, start_time)

        # --- 4. 结果还原 ---
        deduped_results = []
        seen_hashes = set()

        for result in raw_results:
            processed_res = {}
            
            # 初始化结构
            for d_key in debits.keys():
                processed_res[d_key] = {}

            if is_transposed:
                # 翻转: Driver(贷) -> Bucket(借)
                for credit_acc, d_map in result.items():
                    for debit_acc, amount in d_map.items():
                        if abs(amount) > 0.001:
                            # 借方是被动方(Bucket)，取借方符号
                            d_sign = debit_signs.get(debit_acc, 1)
                            final_amt = amount * d_sign
                            processed_res[debit_acc][credit_acc] = final_amt
            else:
                # 正常: Driver(借) -> Bucket(贷)
                for debit_acc, c_map in result.items():
                    for credit_acc, amount in c_map.items():
                        if abs(amount) > 0.001:
                            # 借方是主动方(Driver)，取借方符号
                            d_sign = debit_signs.get(debit_acc, 1)
                            final_amt = amount * d_sign
                            processed_res[debit_acc][credit_acc] = final_amt

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
        # 排序：从大到小
        driver_items = sorted(drivers_dict.items(), key=lambda x: x[1], reverse=True)
        bucket_items = list(buckets_dict.items()) 
        
        results = []
        is_timeout = [False]

        def find_splits(target_amt, available_buckets):
            valid_splits = []
            n = len(available_buckets)
            
            # A. 全匹配
            for r in range(1, n + 1):
                for indices in itertools.combinations(range(n), r):
                    subset_sum = sum(available_buckets[i][1] for i in indices)
                    if abs(subset_sum - target_amt) < 0.001:
                        split_map = {}
                        for i in indices:
                            split_map[available_buckets[i][0]] = available_buckets[i][1]
                        valid_splits.append(split_map)

            # B. 部分匹配 (至多1个非边界)
            for i in range(n):
                partial_name, partial_cap = available_buckets[i]
                others_indices = [x for x in range(n) if x != i]
                n_others = len(others_indices)
                
                for r in range(n_others + 1):
                    for sub_indices in itertools.combinations(others_indices, r):
                        current_sum = sum(available_buckets[k][1] for k in sub_indices)
                        needed = target_amt - current_sum
                        
                        if 0.001 < needed < partial_cap - 0.001:
                            split_map = {}
                            for k in sub_indices:
                                split_map[available_buckets[k][0]] = available_buckets[k][1]
                            split_map[partial_name] = needed
                            valid_splits.append(split_map)
            return valid_splits

        def dfs(d_idx, current_allocations, current_buckets):
            if len(results) >= max_sol: return
            if time.time() - start_time > timeout:
                is_timeout[0] = True; return

            if d_idx == len(driver_items):
                remain = sum(amt for _, amt in current_buckets)
                if remain < 0.001:
                    results.append(current_allocations)
                return

            driver_name, driver_amt = driver_items[d_idx]
            possible_splits = find_splits(driver_amt, current_buckets)
            
            if not possible_splits: return

            for split in possible_splits:
                if len(results) >= max_sol: return
                
                next_alloc = current_allocations.copy()
                next_alloc[driver_name] = split
                
                next_buckets = []
                for b_name, b_amt in current_buckets:
                    used = split.get(b_name, 0)
                    remain = b_amt - used
                    if remain > 0.001:
                        next_buckets.append((b_name, remain))
                
                dfs(d_idx + 1, next_alloc, next_buckets)

        dfs(0, {}, bucket_items)
        return results, is_timeout[0]