import itertools
import hashlib
import time
from collections import defaultdict

class ExhaustiveSolver:
    """
    穷举算法最终版 (最大资源占优策略)
    特性：
    1. 智能转置：始终用"少"的一方遍历"多"的一方
    2. 最大优先：从大金额开始匹配 (Big Rock First)，强力剪枝
    3. 严格边界：每组拆分至多包含1个非边界值
    4. 超时熔断：防止死循环
    """
    
    @staticmethod
    def calculate_combinations(debit_ledger, credit_ledger, max_solutions=200, timeout=2.0):
        start_time = time.time()
        
        # 1. 预处理 (符号映射 & 过滤0值)
        sign_mapping = {}
        for account, amount in debit_ledger.items():
            sign_mapping[account] = 1 if amount >= 0 else -1
            debit_ledger[account] = abs(amount)

        for account, amount in credit_ledger.items():
            sign_mapping[account] = 1 if amount >= 0 else -1
            credit_ledger[account] = abs(amount)

        debits = {k: v for k, v in debit_ledger.items() if abs(v) > 0.001}
        credits = {k: v for k, v in credit_ledger.items() if abs(v) > 0.001}
        
        # 2. 智能转置判断
        # 此时只看数量，数量少的做 Driver
        n_debits = len(debits)
        n_credits = len(credits)
        is_transposed = False
        
        if n_credits < n_debits and n_credits > 0:
            is_transposed = True
            drivers = credits
            buckets = debits
        else:
            drivers = debits
            buckets = credits

        # 3. 核心计算
        raw_results, is_timeout = ExhaustiveSolver._core_solve(drivers, buckets, max_solutions, timeout, start_time)

        # 4. 结果还原 & 去重
        deduped_results = []
        seen_hashes = set()

        for result in raw_results:
            processed_res = {}
            
            # 初始化结构
            for d_key in debits.keys():
                processed_res[d_key] = {}

            if is_transposed:
                # 翻转回来: Driver(贷) -> Bucket(借)
                for credit_acc, d_map in result.items():
                    for debit_acc, amount in d_map.items():
                        if abs(amount) > 0.001:
                            d_sign = sign_mapping.get(debit_acc, 1)
                            final_amt = amount * d_sign
                            processed_res[debit_acc][credit_acc] = final_amt
            else:
                # 正常: Driver(借) -> Bucket(贷)
                for debit_acc, c_map in result.items():
                    for credit_acc, amount in c_map.items():
                        if abs(amount) > 0.001:
                            d_sign = sign_mapping.get(debit_acc, 1)
                            final_amt = amount * d_sign
                            processed_res[debit_acc][credit_acc] = final_amt

            # 生成唯一签名用于去重
            temp_sig = []
            sorted_d = sorted(processed_res.keys())
            for d in sorted_d:
                sorted_c = sorted(processed_res[d].keys())
                for c in sorted_c:
                    amt = processed_res[d][c]
                    if abs(amt) > 0.001:
                        # round 4 防止浮点微小差异导致去重失败
                        temp_sig.append(f"{d}:{c}:{round(amt, 4)}")
            
            sig_str = "|".join(temp_sig)
            res_hash = hashlib.md5(sig_str.encode()).hexdigest()
            
            if res_hash not in seen_hashes:
                seen_hashes.add(res_hash)
                deduped_results.append(processed_res)

        return deduped_results, is_timeout

    @staticmethod
    def _core_solve(drivers_dict, buckets_dict, max_sol, timeout, start_time):
        """
        drivers: 主动遍历方 (Sorted DESC)
        buckets: 被动分配方
        """
        # === 关键修改：从大到小排序 (Reverse=True) ===
        # 最大资源占优策略：先填大坑
        driver_items = sorted(drivers_dict.items(), key=lambda x: x[1], reverse=True)
        
        # bucket 也可以排个序，虽然 itertools 组合是按位置的，但排个序看着整齐
        bucket_items = sorted(list(buckets_dict.items()), key=lambda x: x[1], reverse=True)
        
        results = []
        is_timeout = [False]

        # 内部生成器
        def find_splits(target_amt, available_buckets):
            valid_splits = []
            n = len(available_buckets)
            
            # 这里的 available_buckets 已经是 list of tuples
            
            # A. 全匹配 (0个非边界)
            # 优先匹配组合数少的？其实从小到大遍历组合长度即可
            for r in range(1, n + 1):
                for indices in itertools.combinations(range(n), r):
                    subset_sum = sum(available_buckets[i][1] for i in indices)
                    
                    if abs(subset_sum - target_amt) < 0.001:
                        split_map = {}
                        for i in indices:
                            b_name, b_amt = available_buckets[i]
                            split_map[b_name] = b_amt
                        valid_splits.append(split_map)

            # B. 部分匹配 (1个非边界)
            for i in range(n):
                partial_name, partial_cap = available_buckets[i]
                
                # 剩余需要的钱由其他 buckets 凑出
                others_indices = [x for x in range(n) if x != i]
                n_others = len(others_indices)
                
                for r in range(n_others + 1):
                    for sub_indices in itertools.combinations(others_indices, r):
                        current_sum = sum(available_buckets[k][1] for k in sub_indices)
                        needed = target_amt - current_sum
                        
                        # 校验 needed 是否合法
                        if 0.001 < needed < partial_cap - 0.001:
                            split_map = {}
                            # 全额部分
                            for k in sub_indices:
                                b_name, b_amt = available_buckets[k]
                                split_map[b_name] = b_amt
                            # 差额部分
                            split_map[partial_name] = needed
                            valid_splits.append(split_map)
            
            return valid_splits

        # DFS
        def dfs(d_idx, current_allocations, current_buckets):
            if len(results) >= max_sol: return
            if time.time() - start_time > timeout:
                is_timeout[0] = True; return

            # 结束条件
            if d_idx == len(driver_items):
                # 检查余额是否清零
                remain = sum(amt for _, amt in current_buckets)
                if remain < 0.001:
                    results.append(current_allocations)
                return

            driver_name, driver_amt = driver_items[d_idx]
            
            # 寻找拆分方案
            possible_splits = find_splits(driver_amt, current_buckets)
            
            # 如果没找到方案，说明此路不通 (因为每个 Driver 必须被填满)
            if not possible_splits:
                return

            for split in possible_splits:
                if len(results) >= max_sol: return
                
                next_alloc = current_allocations.copy()
                next_alloc[driver_name] = split
                
                # 更新 buckets
                next_buckets = []
                for b_name, b_amt in current_buckets:
                    used = split.get(b_name, 0)
                    remain = b_amt - used
                    if remain > 0.001:
                        next_buckets.append((b_name, remain))
                
                dfs(d_idx + 1, next_alloc, next_buckets)

        dfs(0, {}, bucket_items)
        return results, is_timeout[0]