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
        self.df = pd.read_excel(file_path, dtype=str)
        
        date_col = mapping['date']
        voucher_col = mapping['voucher_id']
        summ_col = mapping['summary']
        
        self.df['_uid'] = self.df[date_col].astype(str) + "_" + self.df[voucher_col].astype(str)
        
        self.df['_calc_debit'] = pd.to_numeric(self.df[mapping['debit']], errors='coerce').fillna(0)
        self.df['_calc_credit'] = pd.to_numeric(self.df[mapping['credit']], errors='coerce').fillna(0)
        
        subj_col = mapping['subject']
        self.df['_calc_subj'] = self.df[subj_col].astype(str).str.strip()

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
            
            # 1. 保守清洗
            clean_group = self._get_cleaned_amounts(group)
            
            debits = clean_group[clean_group['_calc_debit'].abs() > 0.0001]
            credits = clean_group[clean_group['_calc_credit'].abs() > 0.0001]
            
            unique_subjs = set(clean_group['_calc_subj'])
            if "本年利润" in unique_subjs:
                processed_count += 1
                simple_count += 1
                continue

            if debits.empty and credits.empty: continue

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
        保守清洗：只有单边畸形才移动负数，双边正常数据绝不触碰。
        """
        df = group.copy()
        has_debit_values = (df['_calc_debit'].abs() > 0.001).any()
        has_credit_values = (df['_calc_credit'].abs() > 0.001).any()

        # 双边都有数据 -> 保持原样 (Case: 你的表1)
        if has_debit_values and has_credit_values:
            return df

        # 单边清洗
        if has_debit_values and not has_credit_values:
            mask_d_neg = df['_calc_debit'] < 0
            df.loc[mask_d_neg, '_calc_credit'] = df.loc[mask_d_neg, '_calc_debit'].abs()
            df.loc[mask_d_neg, '_calc_debit'] = 0
            
        elif has_credit_values and not has_debit_values:
            mask_c_neg = df['_calc_credit'] < 0
            df.loc[mask_c_neg, '_calc_debit'] = df.loc[mask_c_neg, '_calc_credit'].abs()
            df.loc[mask_c_neg, '_calc_credit'] = 0
            
        return df

    def _add_to_cluster(self, uid, debits, credits):
        d_subjs = sorted(debits['_calc_subj'].tolist())
        c_subjs = sorted(credits['_calc_subj'].tolist())
        key_str = "、".join(sorted(list(set(d_subjs) | set(c_subjs))))
        key_hash = hashlib.md5(key_str.encode()).hexdigest()
        
        self.complex_clusters[key_hash].append(uid)
        
        d_dict = defaultdict(float)
        for _, row in debits.iterrows(): d_dict[row['_calc_subj']] += row['_calc_debit']
        c_dict = defaultdict(float)
        for _, row in credits.iterrows(): c_dict[row['_calc_subj']] += row['_calc_credit']
        
        self.complex_data_cache[uid] = {"debits": dict(d_dict), "credits": dict(c_dict)}

        if key_hash not in self.cluster_samples:
            self.cluster_samples[key_hash] = {
                "name": key_str, "debits": dict(d_dict), "credits": dict(c_dict),
                "count": 1, "sample_uid": uid
            }
        else:
            self.cluster_samples[key_hash]["count"] += 1

    def finalize_report(self, kb, log_callback):
        solver = ExhaustiveSolver()
        final_rows = []
        
        grouped = self.df.groupby('_uid', sort=False)
        total_groups = len(grouped)
        processed = 0
        
        original_cols = [c for c in self.df.columns if not c.startswith('_')]
        
        for uid, group in grouped:
            processed += 1
            if processed % 100 == 0: log_callback(f"生成进度: {processed}/{total_groups}...")
            
            clean_group = self._get_cleaned_amounts(group)
            
            debits = clean_group[clean_group['_calc_debit'].abs() > 0.0001]
            credits = clean_group[clean_group['_calc_credit'].abs() > 0.0001]
            
            if debits.empty and credits.empty:
                self._append_original_rows(final_rows, group, original_cols, "无效分录")
                continue

            unique_subjs = set(clean_group['_calc_subj'])
            if "本年利润" in unique_subjs:
                self._append_closing_entry(final_rows, group, original_cols)
                continue

            d_types = len(set(debits['_calc_subj']))
            c_types = len(set(credits['_calc_subj']))

            if (d_types == 1 and c_types == 1) or \
               (d_types == 1 and c_types > 1) or \
               (d_types > 1 and c_types == 1):
                if d_types == 1 and c_types == 1:
                    target_c = credits.iloc[0]['_calc_subj']
                    target_d = debits.iloc[0]['_calc_subj']
                    self._append_simple_rows(final_rows, group, original_cols, target_c, target_d)
                else:
                    self._append_1vN_rows_reconstruct(final_rows, uid, original_cols, debits, credits, d_types==1)
            else:
                self._append_complex_rows(final_rows, group, original_cols, uid, kb, solver)

        df_final = pd.DataFrame(final_rows)
        final_cols = original_cols + ["对方科目"]
        for c in final_cols:
            if c not in df_final.columns: df_final[c] = ""
            
        return df_final[final_cols]

    # --- 辅助函数 ---
    def _copy_row_data(self, row, cols):
        return {c: row[c] for c in cols}

    def _create_virtual_row(self, uid, cols, subj, debit_amt, credit_amt, contra):
        meta = self.meta_cache.get(uid, {})
        row = {}
        row[self.mapping['date']] = meta.get('date', '')
        row[self.mapping['voucher_id']] = meta.get('voucher_id', '')
        row[self.mapping['summary']] = meta.get('summary', '') 
        
        row[self.mapping['subject']] = subj
        for c in cols:
            if c not in row: row[c] = ""
            
        row[self.mapping['debit']] = debit_amt if debit_amt is not None else None
        row[self.mapping['credit']] = credit_amt if credit_amt is not None else None
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
            if abs(row['_calc_debit']) > 0.001: new_row["对方科目"] = target_c
            elif abs(row['_calc_credit']) > 0.001: new_row["对方科目"] = target_d
            final_rows.append(new_row)

    def _append_1vN_rows_reconstruct(self, final_rows, uid, cols, debits, credits, is_1_debit):
        single_side_subj = debits.iloc[0]['_calc_subj'] if is_1_debit else credits.iloc[0]['_calc_subj']
        multi_side_rows = credits if is_1_debit else debits
        
        for _, row in multi_side_rows.iterrows():
            row_multi = self._copy_row_data(row, cols)
            row_multi["对方科目"] = single_side_subj
            final_rows.append(row_multi)
            
            amount = row['_calc_credit'] if is_1_debit else row['_calc_debit']
            if is_1_debit:
                row_single = self._create_virtual_row(uid, cols, single_side_subj, amount, None, row['_calc_subj'])
            else:
                row_single = self._create_virtual_row(uid, cols, single_side_subj, None, amount, row['_calc_subj'])
            final_rows.append(row_single)

    def _append_complex_rows(self, final_rows, group, cols, uid, kb, solver):
        data = self.complex_data_cache.get(uid)
        if not data:
            self._append_original_rows(final_rows, group, cols, "缓存丢失")
            return

        solutions, _ = solver.calculate_combinations(data['debits'], data['credits'], max_solutions=50, timeout=0.5)
        
        if not solutions:
            self._append_original_rows(final_rows, group, cols, "需人工分析(无解)")
            return

        ranked = kb.rank_solutions(solutions)
        best_sol = ranked[0] 
        
        # 借方行
        for d_subj, c_map in best_sol.items():
            for c_subj, amt in c_map.items():
                if abs(amt) > 0.001:
                    # 注意：这里的 amt 带有借方符号
                    row = self._create_virtual_row(uid, cols, d_subj, amt, None, c_subj)
                    final_rows.append(row)
        
        # 贷方行
        c_side_map = defaultdict(dict)
        for d_subj, c_map in best_sol.items():
            for c_subj, amt in c_map.items():
                if abs(amt) > 0.001:
                    c_side_map[c_subj][d_subj] = amt
        
        for c_subj, d_map in c_side_map.items():
            for d_subj, amt in d_map.items():
                # 贷方行需要贷方符号，通常 amt 此时是借方符号
                # 但在 _create_virtual_row 中我们是放入 credit_amt 参数
                # 所以要取绝对值放入，或者保持符号？
                # 一般贷方列金额是正数，表示贷方发生。如果是红字冲销，是负数。
                # 这里的 logic 略微复杂，为了稳妥，我们通过 group 原始数据来校验符号？
                # 不，直接用算法返回的 signed amount 即可。
                # 如果是正常的借100，算法返回100。放入 Credit 列，就是贷100。
                # 如果是红字借-100，算法返回-100。放入 Credit 列，就是贷-100 (即借100)。
                # 逻辑是通的。
                row = self._create_virtual_row(uid, cols, c_subj, None, amt, d_subj)
                final_rows.append(row)