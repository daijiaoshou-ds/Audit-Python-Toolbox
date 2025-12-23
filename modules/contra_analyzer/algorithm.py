import itertools
import hashlib
import time
from collections import defaultdict

class ExhaustiveSolver:
    """
    穷举算法 v8.0 (资金科目特别保护版)
    特性：
    1. 资金保护：'银行'/'现金' 科目禁止参与正负抵消，防止发生额虚增。
    2. 智能转置：多遍历少。
    3. 最小优先：按代数大小排序。
    """
    
    # 定义敏感科目关键词
    SENSITIVE_KEYWORDS = ["银行", "现金", "Bank", "Cash"]

    @staticmethod
    def is_sensitive(key_name):
        """检查是否为资金类科目"""
        return any(kw in key_name for kw in ExhaustiveSolver.SENSITIVE_KEYWORDS)

    @staticmethod
    def calculate_combinations(debit_ledger, credit_ledger, max_solutions=200, timeout=5.0):
        start_time = time.time()
        
        # 1. 预处理
        debits = {k: v for k, v in debit_ledger.items() if abs(v) > 0.001}
        credits = {k: v for k, v in credit_ledger.items() if abs(v) > 0.001}
        
        # 2. 智能转置
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

        # 4. 结果还原 & 去重 & 会计校验
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
            
            # 后置校验：确保金额守恒
            is_valid_accounting = True
            # 校验借方
            for d_key, original_amt in debits.items():
                split_sum = sum(processed_res[d_key].values())
                if abs(split_sum - original_amt) > 0.001:
                    is_valid_accounting = False; break
            if not is_valid_accounting: continue
            
            # 校验贷方
            c_received = defaultdict(float)
            for d_key, c_map in processed_res.items():
                for c_key, amt in c_map.items():
                    c_received[c_key] += amt
            for c_key, original_amt in credits.items():
                if abs(c_received[c_key] - original_amt) > 0.001:
                    is_valid_accounting = False; break
            if not is_valid_accounting: continue

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
        # 排序：代数值从小到大
        driver_items = sorted(drivers_dict.items(), key=lambda x: x[1], reverse=False)
        bucket_items = sorted(list(buckets_dict.items()), key=lambda x: x[1], reverse=False)
        
        results = []
        is_timeout = [False]

        def generate_combinations(driver_name, target_amt, available_buckets):
            valid_splits = []
            n = len(available_buckets)
            seen = set()
            
            # 检查 Driver 是否敏感 (如银行存款)
            is_driver_sensitive = ExhaustiveSolver.is_sensitive(driver_name)
            
            # A. 全匹配 (Subset Sum)
            for r in range(1, n + 1):
                for indices in itertools.combinations(range(n), r):
                    subset_sum = sum(available_buckets[i][1] for i in indices)
                    
                    if abs(subset_sum - target_amt) < 0.001:
                        
                        # [保护] 如果 Driver 是敏感科目，严禁其拆分对象包含异号 (防止虚增)
                        if is_driver_sensitive:
                            has_mixed_sign = False
                            for i in indices:
                                # 如果 Driver是正，Bucket必须非负；Driver是负，Bucket必须非正
                                b_amt = available_buckets[i][1]
                                if target_amt > 0 and b_amt < -0.001: has_mixed_sign = True
                                if target_amt < 0 and b_amt > 0.001: has_mixed_sign = True
                            if has_mixed_sign: continue

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
                
                # [保护] 如果 Bucket 是敏感科目 (如银行存款)
                # 它只能接受"包含"关系的拆分，严禁"扩容"
                is_bucket_sensitive = ExhaustiveSolver.is_sensitive(partial_name)

                others_indices = [x for x in range(n) if x != i]
                n_others = len(others_indices)
                
                for r in range(n_others + 1):
                    for sub_indices in itertools.combinations(others_indices, r):
                        current_sum = sum(available_buckets[k][1] for k in sub_indices)
                        needed = target_amt - current_sum
                        
                        if abs(needed) < 0.001: continue
                        
                        # --- 核心判断逻辑 ---
                        is_valid_part = False
                        
                        if is_bucket_sensitive:
                            # 严格模式：Needed 必须在 Partial_Cap 内部
                            # 同号，且绝对值更小
                            if partial_cap > 0:
                                if 0.001 < needed < partial_cap - 0.001: is_valid_part = True
                            elif partial_cap < 0:
                                if partial_cap + 0.001 < needed < -0.001: is_valid_part = True
                        else:
                            # 宽松模式：只要不等于 Cap 即可 (允许扩容/反向)
                            if abs(needed - partial_cap) > 0.001:
                                is_valid_part = True
                                
                        if is_valid_part:
                            # [保护] 如果 Driver 是敏感科目，再次检查 sub_indices 的符号
                            if is_driver_sensitive:
                                # needed 已经通过上面的检查，现在检查 others
                                has_mixed_sign = False
                                # 检查 others
                                for k in sub_indices:
                                    b_amt = available_buckets[k][1]
                                    if target_amt > 0 and b_amt < -0.001: has_mixed_sign = True
                                    if target_amt < 0 and b_amt > 0.001: has_mixed_sign = True
                                # 检查 needed (needed也是组成部分)
                                if target_amt > 0 and needed < -0.001: has_mixed_sign = True
                                if target_amt < 0 and needed > 0.001: has_mixed_sign = True
                                
                                if has_mixed_sign: continue

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
            if len(results) >= max_sol * 2: return 
            if time.time() - start_time > timeout:
                is_timeout[0] = True; return

            if d_idx == len(driver_items):
                remain = sum(amt for _, amt in current_buckets)
                if abs(remain) < 0.01:
                    results.append(current_allocations)
                return

            driver_name, driver_amt = driver_items[d_idx]
            
            # 传入 driver_name 以便检查敏感性
            possible_splits = generate_combinations(driver_name, driver_amt, current_buckets)
            
            if not possible_splits: return

            for split in possible_splits:
                if len(results) >= max_sol * 2: return
                
                next_alloc = current_allocations.copy()
                next_alloc[driver_name] = split
                
                next_buckets = []
                for b_name, b_amt in current_buckets:
                    used = split.get(b_name, 0)
                    remain = b_amt - used
                    if abs(remain) > 0.001:
                        next_buckets.append((b_name, remain))
                
                dfs(d_idx + 1, next_alloc, next_buckets)

        dfs(0, {}, bucket_items)
        return results, is_timeout[0]