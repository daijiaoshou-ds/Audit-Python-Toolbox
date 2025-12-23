import pandas as pd
import hashlib
from collections import defaultdict
from .algorithm import ExhaustiveSolver

class ContraProcessor:
    def __init__(self):
        self.df = None
        self.mapping = {} 
        self.complex_data_cache = {}
        self.meta_cache = {} 

    def load_data(self, file_path, mapping):
        self.mapping = mapping
        # 1. 读取原始数据
        self.df = pd.read_excel(file_path, dtype=str)
        
        date_col = mapping['date']
        voucher_col = mapping['voucher_id']
        summ_col = mapping['summary']
        
        # 2. 生成唯一ID
        self.df['_uid'] = self.df[date_col].astype(str) + "_" + self.df[voucher_col].astype(str)
        
        # 3. 转换金额
        self.df['_calc_debit'] = pd.to_numeric(self.df[mapping['debit']], errors='coerce').fillna(0)
        self.df['_calc_credit'] = pd.to_numeric(self.df[mapping['credit']], errors='coerce').fillna(0)
        
        # 4. 科目去空格
        subj_col = mapping['subject']
        self.df['_calc_subj'] = self.df[subj_col].astype(str).str.strip()

        # 5. 缓存元数据
        for uid, group in self.df.groupby('_uid'):
            first_row = group.iloc[0]
            unique_summs = group[summ_col].dropna().unique()
            combined_summ = " | ".join([str(s) for s in unique_summs if str(s).strip()])
            
            self.meta_cache[uid] = {
                'date': first_row[date_col],
                'voucher_id': first_row[voucher_col],
                'summary': combined_summ
            }

    def process_all(self, stop_event=None):
        self.complex_clusters = defaultdict(list)
        self.cluster_samples = {}
        self.complex_data_cache = {}
        
        grouped = self.df.groupby('_uid')
        processed_count = 0
        simple_count = 0

        for uid, group in grouped:
            if stop_event and stop_event.is_set(): break
            
            # 1. 智能清洗
            clean_group = self._get_cleaned_amounts(group)
            
            # 2. 提取有效行
            debits = clean_group[clean_group['_calc_debit'].abs() > 0.0001]
            credits = clean_group[clean_group['_calc_credit'].abs() > 0.0001]
            
            # 3. 损益拦截
            unique_subjs = set(clean_group['_calc_subj'])
            if "本年利润" in unique_subjs:
                processed_count += 1
                simple_count += 1
                continue

            if debits.empty and credits.empty: continue

            # 4. 分层逻辑
            d_types = len(set(debits['_calc_subj']))
            c_types = len(set(credits['_calc_subj']))

            if (d_types == 1 and c_types == 1) or \
               (d_types == 1 and c_types > 1) or \
               (d_types > 1 and c_types == 1):
                simple_count += 1
            else:
                self._add_to_cluster(uid, debits, credits)
            
            processed_count += 1

        return {
            "processed": processed_count,
            "complex_groups": len(self.complex_clusters),
            "simple_solved": simple_count
        }

    def _get_cleaned_amounts(self, group):
        """
        修正后的清洗逻辑：
        判断是否为【全借】或【全贷】。
        只有在这两种情况下，才执行"负数搬家"。
        """
        df = group.copy()
        
        # 使用 abs().sum() 判断该列是否有值
        total_debit_activity = df['_calc_debit'].abs().sum()
        total_credit_activity = df['_calc_credit'].abs().sum()

        is_all_debit = (total_credit_activity < 0.001) and (total_debit_activity > 0.001)
        is_all_credit = (total_debit_activity < 0.001) and (total_credit_activity > 0.001)

        if is_all_debit:
            # 全借分录：将借方负数 -> 移至贷方(变正)
            mask_d_neg = df['_calc_debit'] < 0
            if mask_d_neg.any():
                df.loc[mask_d_neg, '_calc_credit'] = df.loc[mask_d_neg, '_calc_debit'].abs()
                df.loc[mask_d_neg, '_calc_debit'] = 0
            
        elif is_all_credit:
            # 全贷分录：将贷方负数 -> 移至借方(变正)
            mask_c_neg = df['_calc_credit'] < 0
            if mask_c_neg.any():
                df.loc[mask_c_neg, '_calc_debit'] = df.loc[mask_c_neg, '_calc_credit'].abs()
                df.loc[mask_c_neg, '_calc_credit'] = 0
            
        return df

    def _add_to_cluster(self, uid, debits, credits):
        d_subjs = sorted(debits['_calc_subj'].tolist())
        c_subjs = sorted(credits['_calc_subj'].tolist())
        
        key_str = "、".join(sorted(list(set(d_subjs) | set(c_subjs))))
        key_hash = hashlib.md5(key_str.encode()).hexdigest()
        
        self.complex_clusters[key_hash].append(uid)
        
        # 聚合逻辑：保留方向后缀
        d_dict = defaultdict(float)
        for subj, amt in zip(debits['_calc_subj'], debits['_calc_debit']):
            # 全借分录清洗后，原来的负借变成了正贷，所以这里进来的都是正数(或原本的正借)
            # 但如果是双边分录，这里可能进负数。
            # 为了防止同名抵消，必须区分正负。
            suffix = "Pos" if amt >= 0 else "Neg"
            key = f"{subj}__{suffix}__D"
            d_dict[key] += amt
            
        c_dict = defaultdict(float)
        for subj, amt in zip(credits['_calc_subj'], credits['_calc_credit']):
            suffix = "Pos" if amt >= 0 else "Neg"
            key = f"{subj}__{suffix}__C"
            c_dict[key] += amt
        
        self.complex_data_cache[uid] = {"debits": dict(d_dict), "credits": dict(c_dict)}

        if key_hash not in self.cluster_samples:
            self.cluster_samples[key_hash] = {
                "name": key_str, 
                "debits": dict(d_dict), 
                "credits": dict(c_dict),
                "count": 1, 
                "sample_uid": uid
            }
        else:
            self.cluster_samples[key_hash]["count"] += 1

    def finalize_report(self, kb, log_callback):
            solver = ExhaustiveSolver()
            final_rows = []
            
            grouped = self.df.groupby('_uid', sort=False)
            total_groups = len(grouped)
            processed = 0
            
            # 原始列名
            original_cols = [c for c in self.df.columns if not c.startswith('_')]
            
            for uid, group in grouped:
                processed += 1
                if processed % 100 == 0: log_callback(f"生成进度: {processed}/{total_groups}...")
                
                clean_group = self._get_cleaned_amounts(group)
                debits = clean_group[clean_group['_calc_debit'].abs() > 0.0001]
                credits = clean_group[clean_group['_calc_credit'].abs() > 0.0001]
                
                if debits.empty and credits.empty:
                    self._append_original_rows(final_rows, group, original_cols, "无效分录"); continue

                unique_subjs = set(clean_group['_calc_subj'])
                if "本年利润" in unique_subjs:
                    self._append_closing_entry(final_rows, group, original_cols); continue

                d_types = len(set(debits['_calc_subj']))
                c_types = len(set(credits['_calc_subj']))

                if (d_types == 1 and c_types == 1) or \
                (d_types == 1 and c_types > 1) or \
                (d_types > 1 and c_types == 1):
                    if d_types == 1 and c_types == 1:
                        self._append_simple_rows(final_rows, group, original_cols, credits.iloc[0]['_calc_subj'], debits.iloc[0]['_calc_subj'])
                    else:
                        self._append_1vN_rows_reconstruct(final_rows, uid, original_cols, debits, credits, d_types==1)
                else:
                    self._append_complex_rows(final_rows, group, original_cols, uid, kb, solver)

            df_final = pd.DataFrame(final_rows)
            
            # === 核心修改：列重排 ===
            # 目标：将 "对方科目" 插入到 "贷方金额" 后面
            # 如果找不到贷方列，就放到最后
            
            output_cols = []
            credit_col_name = self.mapping['credit']
            
            if credit_col_name in original_cols:
                idx = original_cols.index(credit_col_name)
                # 插入到贷方后面
                output_cols = original_cols[:idx+1] + ["对方科目"] + original_cols[idx+1:]
            else:
                output_cols = original_cols + ["对方科目"]
                
            # 补全缺失列
            for c in output_cols:
                if c not in df_final.columns: df_final[c] = ""
                
            # 统一转数值
            df_final[self.mapping['debit']] = pd.to_numeric(df_final[self.mapping['debit']], errors='coerce').fillna(0)
            df_final[self.mapping['credit']] = pd.to_numeric(df_final[self.mapping['credit']], errors='coerce').fillna(0)
            
            return df_final[output_cols]

    # --- 辅助函数 ---
    
    def _copy_row_data(self, row, cols):
        """复制原始列，但忽略金额列(因为金额可能被清洗修改过)"""
        # 注意：这里我们只复制非金额列。金额列将在后续步骤中由 _calc_debit/credit 填充
        # 或者：我们直接复制，然后在外部覆盖金额
        return {c: row[c] for c in cols}

    def _create_virtual_row(self, uid, cols, subj, debit_amt, credit_amt, contra):
        meta = self.meta_cache.get(uid, {})
        row = {}
        row[self.mapping['date']] = meta.get('date', '')
        row[self.mapping['voucher_id']] = meta.get('voucher_id', '')
        row[self.mapping['summary']] = meta.get('summary', '') 
        row[self.mapping['subject']] = subj
        
        # 填充其他列为空
        for c in cols:
            if c not in row: row[c] = ""
            
        row[self.mapping['debit']] = debit_amt if debit_amt is not None else 0
        row[self.mapping['credit']] = credit_amt if credit_amt is not None else 0
        row["对方科目"] = contra
        return row

    def _append_original_rows(self, final_rows, group, cols, contra_msg):
        for _, row in group.iterrows():
            new_row = self._copy_row_data(row, cols)
            new_row["对方科目"] = contra_msg
            final_rows.append(new_row)

    def _append_closing_entry(self, final_rows, group, cols):
        for _, row in group.iterrows():
            new_row = self._copy_row_data(row, cols)
            if abs(row['_calc_debit']) > 0.001 or abs(row['_calc_credit']) > 0.001:
                new_row["对方科目"] = "本年利润"
            final_rows.append(new_row)

    def _append_simple_rows(self, final_rows, group, cols, target_c, target_d):
        for _, row in group.iterrows():
            new_row = self._copy_row_data(row, cols)
            # 使用清洗后的金额列
            if abs(row['_calc_debit']) > 0.001: 
                new_row[self.mapping['debit']] = row['_calc_debit']
                new_row[self.mapping['credit']] = 0
                new_row["对方科目"] = target_c
            elif abs(row['_calc_credit']) > 0.001: 
                new_row[self.mapping['credit']] = row['_calc_credit']
                new_row[self.mapping['debit']] = 0
                new_row["对方科目"] = target_d
            final_rows.append(new_row)

    def _append_1vN_rows_reconstruct(self, final_rows, uid, cols, debits, credits, is_1_debit):
        single_side_subj = debits.iloc[0]['_calc_subj'] if is_1_debit else credits.iloc[0]['_calc_subj']
        multi_side_rows = credits if is_1_debit else debits
        
        for _, row in multi_side_rows.iterrows():
            # 1. 保留 N 的这一行
            row_multi = self._copy_row_data(row, cols)
            
            # 关键：覆盖金额为清洗后的金额 (正数)
            if is_1_debit:
                row_multi[self.mapping['credit']] = row['_calc_credit']
                row_multi[self.mapping['debit']] = 0
            else:
                row_multi[self.mapping['debit']] = row['_calc_debit']
                row_multi[self.mapping['credit']] = 0
                
            row_multi["对方科目"] = single_side_subj
            final_rows.append(row_multi)
            
            # 2. 生成对应的 1 的拆分行
            amount = row['_calc_credit'] if is_1_debit else row['_calc_debit']
            if is_1_debit:
                row_single = self._create_virtual_row(uid, cols, single_side_subj, amount, None, row['_calc_subj'])
            else:
                row_single = self._create_virtual_row(uid, cols, single_side_subj, None, amount, row['_calc_subj'])
            final_rows.append(row_single)

    def _append_complex_rows(self, final_rows, group, cols, uid, kb, solver):
        """
        group: 必须是 clean_group (负数已挪位)
        """
        data = self.complex_data_cache.get(uid)
        if not data:
            self._append_original_rows(final_rows, group, cols, "缓存丢失")
            return

        solutions, _ = solver.calculate_combinations(data['debits'], data['credits'], max_solutions=200, timeout=1.5)
        
        if not solutions:
            self._append_original_rows(final_rows, group, cols, "需人工分析(无解)")
            return

        ranked = kb.rank_solutions(solutions)
        best_sol = ranked[0] 
        
        # === 借方行重构 ===
        for d_key, c_map in best_sol.items():
            d_subj = d_key.split('__')[0]
            # 找到匹配方向的行 (Pos/Neg)
            target_sign = 1 if "Pos" in d_key else -1
            
            # 这里的 group 是清洗后的，如果是全借分录清洗来的，原来的负借已经变成正贷了。
            # 所以我们要在 debits (_calc_debit) 里找 d_subj。
            
            d_rows = []
            for _, row in group.iterrows():
                if row['_calc_subj'] == d_subj:
                    amt = row['_calc_debit']
                    if abs(amt) > 0.001:
                        # 检查符号: 这里的 amt 已经是清洗后的，如果是全借清洗，这里是正数。
                        # 如果是双边保留，这里可能是负数。
                        # 我们的 d_key (Pos/Neg) 是基于 process_all 时生成的。
                        # 如果 process_all 调用了 clean，那 cache 里的也是 clean 后的。
                        # 所以这里直接比较即可。
                        if (amt >= 0 and target_sign == 1) or (amt < 0 and target_sign == -1):
                            d_rows.append(row)

            total_alloc = sum(c_map.values())
            
            for row in d_rows:
                row_amt = row['_calc_debit']
                if abs(total_alloc) < 0.001: continue
                
                for c_key, target_amt in c_map.items():
                    c_subj = c_key.split('__')[0]
                    if abs(target_amt) > 0.001:
                        ratio = target_amt / total_alloc
                        split_amt = row_amt * ratio
                        
                        new_row = self._copy_row_data(row, cols)
                        new_row[self.mapping['debit']] = split_amt
                        new_row[self.mapping['credit']] = 0 # 确保干净
                        new_row["对方科目"] = c_subj
                        final_rows.append(new_row)

        # === 贷方行重构 ===
        c_side_map = defaultdict(dict)
        for d_key, c_map in best_sol.items():
            for c_key, amt in c_map.items():
                if abs(amt) > 0.001: c_side_map[c_key][d_key] = amt
        
        for c_key, d_map in c_side_map.items():
            c_subj = c_key.split('__')[0]
            target_sign = 1 if "Pos" in c_key else -1
            
            c_rows = []
            for _, row in group.iterrows():
                if row['_calc_subj'] == c_subj:
                    amt = row['_calc_credit']
                    if abs(amt) > 0.001:
                        if (amt >= 0 and target_sign == 1) or (amt < 0 and target_sign == -1):
                            c_rows.append(row)
            
            total_alloc = sum(d_map.values())
            
            for row in c_rows:
                row_amt = row['_calc_credit']
                if abs(total_alloc) < 0.001: continue
                
                for d_key, alloc_amt in d_map.items():
                    d_subj = d_key.split('__')[0]
                    if abs(alloc_amt) > 0.001:
                        ratio = alloc_amt / total_alloc
                        split_amt = row_amt * ratio
                        
                        new_row = self._copy_row_data(row, cols)
                        new_row[self.mapping['credit']] = split_amt
                        new_row[self.mapping['debit']] = 0 # 确保干净
                        new_row["对方科目"] = d_subj
                        final_rows.append(new_row)