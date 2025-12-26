import itertools
import hashlib
import time
from collections import defaultdict

class ExhaustiveSolver:
    """
    穷举算法 v9.1 (先全量后锁定版)
    策略调整：
    1. 优先执行【全量穷举】，保证不漏掉复杂组合。
    2. 随后执行【完美锁定】，作为快速补充和保底。
    3. 结果合并去重。
    """
    
    # 定义敏感科目关键词
    SENSITIVE_KEYWORDS = ["银行", "现金", "Bank", "Cash", "支付宝", "微信"]

    @staticmethod
    def is_sensitive(key_name):
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

        # === 3. 双轨计算 (顺序调整) ===
        
        # 轨道 B: 全量穷举 (主力，不做预锁定)
        # 分配 70% 的时间预算，确保深度搜索
        results_b, timeout_b = ExhaustiveSolver._core_solve(
            drivers, buckets, max_solutions, timeout * 0.7, start_time, use_perfect_lock=False
        )
        
        # 轨道 A: 完美锁定 (快速补充)
        # 利用剩余时间，尝试寻找"捷径解"
        # 即使全量穷举超时了，这个快速版也能瞬间给出几个保底解
        remaining_time = timeout - (time.time() - start_time)
        results_a = []
        
        if remaining_time > 0.1: # 只要还有一点时间
             # 这里 max_solutions 可以传小点，因为只是补充
            results_a, _ = ExhaustiveSolver._core_solve(
                drivers, buckets, max_solutions, remaining_time, time.time(), use_perfect_lock=True
            )

        # 合并结果 (全量在前，锁定在后)
        # 注意：后续的去重逻辑会处理重复项
        raw_results = results_b + results_a
        
        # 标记超时：只要全量版超时了，就标记超时
        is_timeout = timeout_b

        # === 4. 结果还原 & 去重 & 会计校验 ===
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
            for d_key, original_amt in debits.items():
                split_sum = sum(processed_res[d_key].values())
                if abs(split_sum - original_amt) > 0.001:
                    is_valid_accounting = False; break
            if not is_valid_accounting: continue
            
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
                
            if len(deduped_results) >= max_solutions: break

        return deduped_results, is_timeout

    @staticmethod
    def _core_solve(drivers_dict, buckets_dict, max_sol, timeout, start_time, use_perfect_lock=True):
        """
        核心求解器
        """
        driver_items = sorted(drivers_dict.items(), key=lambda x: x[1], reverse=False)
        bucket_items = sorted(list(buckets_dict.items()), key=lambda x: x[1], reverse=False)
        
        # === 预处理：完美锁定 ===
        locked_allocations = {}
        
        if use_perfect_lock:
            temp_buckets = buckets_dict.copy()
            temp_drivers = drivers_dict.copy()
            
            sorted_d_keys = sorted(temp_drivers.keys(), key=lambda k: temp_drivers[k])
            
            for d_key in sorted_d_keys:
                d_val = temp_drivers[d_key]
                found_b = None
                
                for b_key, b_val in temp_buckets.items():
                    if abs(d_val - b_val) < 0.001:
                        # 敏感性检查
                        is_d_sens = ExhaustiveSolver.is_sensitive(d_key)
                        is_b_sens = ExhaustiveSolver.is_sensitive(b_key)
                        if (is_d_sens or is_b_sens) and (d_val * b_val < 0):
                            continue 
                        
                        found_b = b_key
                        break
                
                if found_b:
                    locked_allocations[d_key] = {found_b: d_val}
                    del temp_drivers[d_key]
                    del temp_buckets[found_b]
            
            driver_items = sorted(temp_drivers.items(), key=lambda x: x[1], reverse=False)
            bucket_items = sorted(list(temp_buckets.items()), key=lambda x: x[1], reverse=False)
        
        results = []
        is_timeout = [False]

        def generate_combinations(driver_name, target_amt, available_buckets):
            valid_splits = []
            n = len(available_buckets)
            seen = set()
            
            is_driver_sensitive = ExhaustiveSolver.is_sensitive(driver_name)
            
            # A. 全匹配
            for r in range(1, n + 1):
                for indices in itertools.combinations(range(n), r):
                    subset_sum = sum(available_buckets[i][1] for i in indices)
                    if abs(subset_sum - target_amt) < 0.001:
                        if is_driver_sensitive:
                            has_mixed = False
                            for i in indices:
                                b = available_buckets[i][1]
                                if (target_amt>0 and b<-0.001) or (target_amt<0 and b>0.001): has_mixed=True
                            if has_mixed: continue

                        split_map = {}
                        for i in indices:
                            split_map[available_buckets[i][0]] = available_buckets[i][1]
                        
                        key = tuple(sorted((k, round(v, 4)) for k,v in split_map.items()))
                        if key not in seen:
                            seen.add(key)
                            valid_splits.append(split_map)

            # B. 部分匹配
            for i in range(n):
                partial_name, partial_cap = available_buckets[i]
                is_bucket_sensitive = ExhaustiveSolver.is_sensitive(partial_name)
                
                others_indices = [x for x in range(n) if x != i]
                n_others = len(others_indices)
                
                for r in range(n_others + 1):
                    for sub_indices in itertools.combinations(others_indices, r):
                        current_sum = sum(available_buckets[k][1] for k in sub_indices)
                        needed = target_amt - current_sum
                        
                        if abs(needed) < 0.001: continue
                        
                        is_valid_part = False
                        if is_bucket_sensitive:
                            if partial_cap > 0:
                                if 0.001 < needed < partial_cap - 0.001: is_valid_part = True
                            elif partial_cap < 0:
                                if partial_cap + 0.001 < needed < -0.001: is_valid_part = True
                        else:
                            if abs(needed - partial_cap) > 0.001:
                                is_valid_part = True
                                
                        if is_valid_part:
                            if is_driver_sensitive:
                                has_mixed = False
                                for k in sub_indices:
                                    b = available_buckets[k][1]
                                    if (target_amt>0 and b<-0.001) or (target_amt<0 and b>0.001): has_mixed=True
                                if (target_amt>0 and needed<-0.001) or (target_amt<0 and needed>0.001): has_mixed=True
                                if has_mixed: continue

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
                    final_comb = current_allocations.copy()
                    if use_perfect_lock and locked_allocations:
                        final_comb.update(locked_allocations)
                    results.append(final_comb)
                return

            driver_name, driver_amt = driver_items[d_idx]
            
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

        if not driver_items:
            if locked_allocations: results.append(locked_allocations)
        else:
            dfs(0, {}, bucket_items)
            
        return results, is_timeout[0]