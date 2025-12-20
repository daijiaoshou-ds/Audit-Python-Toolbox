import pandas as pd
import hashlib
from collections import defaultdict

class ContraProcessor:
    def __init__(self):
        self.df = None
        self.mapping = {} 
        self.results = [] 
        self.complex_clusters = defaultdict(list) 
        self.cluster_samples = {} 

    def load_data(self, file_path, mapping):
        self.mapping = mapping
        # 强制读取为字符串，防止精度丢失
        self.df = pd.read_excel(file_path, dtype=str)
        
        date_col = mapping['date']
        voucher_col = mapping['voucher_id']
        
        # 1. 生成唯一ID
        self.df['_uid'] = self.df[date_col].astype(str) + "_" + self.df[voucher_col].astype(str)
        
        # 2. 数值转换
        self.df[mapping['debit']] = pd.to_numeric(self.df[mapping['debit']], errors='coerce').fillna(0)
        self.df[mapping['credit']] = pd.to_numeric(self.df[mapping['credit']], errors='coerce').fillna(0)
        
        # 3. 科目名称去空格
        subj_col = mapping['subject']
        self.df[subj_col] = self.df[subj_col].astype(str).str.strip()

    def process_all(self, stop_event=None):
        self.results = []
        self.complex_clusters = defaultdict(list)
        self.cluster_samples = {}
        
        grouped = self.df.groupby('_uid')
        processed_count = 0

        subj_col = self.mapping['subject']
        d_col = self.mapping['debit']
        c_col = self.mapping['credit']

        for uid, group in grouped:
            if stop_event and stop_event.is_set(): break
            
            # === 1. 智能清洗 (修复负数乱跑问题) ===
            clean_group = self._normalize_entry(group)
            
            # === 2. 损益结转拦截 (新增) ===
            # 如果包含"本年利润"，直接特殊处理，不进算法
            unique_subjs = set(clean_group[subj_col])
            if "本年利润" in unique_subjs:
                self._handle_closing_entry(uid, clean_group)
                processed_count += 1
                continue

            # === 3. 提取借贷方 ===
            # 注意：这里我们不再简单过滤 > 0，而是保留负数，因为现在的算法支持负数
            # 但为了区分方向，我们还是按原始列来分
            debits = clean_group[clean_group[d_col].abs() > 0.0001]
            credits = clean_group[clean_group[c_col].abs() > 0.0001]
            
            if debits.empty and credits.empty: continue

            # === 4. 分层逻辑 ===
            d_types = len(set(debits[subj_col]))
            c_types = len(set(credits[subj_col]))

            # Case 1: 1v1
            if d_types == 1 and c_types == 1:
                self._handle_simple_1v1(uid, debits, credits)
                
            # Case 2: 1vN / Nv1 (一边只有1种科目)
            elif (d_types == 1 and c_types >= 1) or (d_types >= 1 and c_types == 1):
                self._handle_one_side_many(uid, debits, credits, d_types == 1)
                
            # Case 3: NvN (真正的多借多贷)
            else:
                self._add_to_cluster(uid, debits, credits)
            
            processed_count += 1

        return {
            "processed": processed_count,
            "complex_groups": len(self.complex_clusters),
            "simple_solved": len(self.results)
        }

    def _normalize_entry(self, group):
        """
        严谨的负数处理逻辑：
        只有当分录是【完全单边】（如全在借方，含负数）时，才进行借贷置换。
        正常的【有借有贷】（即使含负数），保持原样！
        """
        df = group.copy()
        d_col = self.mapping['debit']
        c_col = self.mapping['credit']
        
        total_d = df[d_col].sum()
        total_c = df[c_col].sum()
        
        has_debit_lines = (df[d_col].abs() > 0.001).any()
        has_credit_lines = (df[c_col].abs() > 0.001).any()

        # 只有在 "只有借方行" 或 "只有贷方行" 的畸形分录下，才启用清洗
        if has_debit_lines and not has_credit_lines:
            # 全是借方 -> 把负数借方移到贷方
            mask_d_neg = df[d_col] < 0
            df.loc[mask_d_neg, c_col] = df.loc[mask_d_neg, d_col].abs()
            df.loc[mask_d_neg, d_col] = 0
            
        elif has_credit_lines and not has_debit_lines:
            # 全是贷方 -> 把负数贷方移到借方
            mask_c_neg = df[c_col] < 0
            df.loc[mask_c_neg, d_col] = df.loc[mask_c_neg, c_col].abs()
            df.loc[mask_c_neg, c_col] = 0
            
        # 正常的有借有贷，哪怕有红字冲销，也保持原样，交给后续逻辑处理
        return df

    def _handle_closing_entry(self, uid, group):
        """处理结转损益分录"""
        subj_col = self.mapping['subject']
        d_col = self.mapping['debit']
        c_col = self.mapping['credit']
        
        # 规则：所有科目的对方科目，强制设为"本年利润"
        # 哪怕是本年利润自己，对方也是本年利润
        
        for _, row in group.iterrows():
            amt = 0
            direction = ""
            if abs(row[d_col]) > 0.001:
                amt = row[d_col]
                direction = "借"
            elif abs(row[c_col]) > 0.001:
                amt = row[c_col]
                direction = "贷"
            else:
                continue

            # 结果直接写入
            self.results.append({
                "uid": uid,
                "src_subject": row[subj_col],
                "contra_subject": "本年利润", # 强制
                "amount": amt,
                "type": "结转损益"
            })

    def _handle_simple_1v1(self, uid, debits, credits):
        target_c = credits.iloc[0][self.mapping['subject']]
        for _, row in debits.iterrows():
            self.results.append({
                "uid": uid,
                "src_subject": row[self.mapping['subject']],
                "contra_subject": target_c,
                "amount": row[self.mapping['debit']],
                "type": "1v1"
            })
        
        # 为了完整性，贷方的对方科目也生成一下 (可选)
        target_d = debits.iloc[0][self.mapping['subject']]
        for _, row in credits.iterrows():
            self.results.append({
                "uid": uid,
                "src_subject": row[self.mapping['subject']],
                "contra_subject": target_d,
                "amount": row[self.mapping['credit']],
                "type": "1v1"
            })

    def _handle_one_side_many(self, uid, debits, credits, is_1_debit):
        """1vN 逻辑"""
        subj_col = self.mapping['subject']
        if is_1_debit:
            # 1借 (A) vs N贷 (B,C)
            # A 的对方是 B和C
            # B,C 的对方是 A
            target_d = debits.iloc[0][subj_col]
            
            # 处理贷方：简单，对方就是 A
            for _, c_row in credits.iterrows():
                self.results.append({
                    "uid": uid,
                    "src_subject": c_row[subj_col],
                    "contra_subject": target_d,
                    "amount": c_row[self.mapping['credit']],
                    "type": "1vN"
                })
                # 同时生成借方的拆分记录: A -> B
                self.results.append({
                    "uid": uid,
                    "src_subject": target_d,
                    "contra_subject": c_row[subj_col],
                    "amount": c_row[self.mapping['credit']],
                    "type": "1vN"
                })
        else:
            # N借 (A,B) vs 1贷 (C)
            target_c = credits.iloc[0][subj_col]
            for _, d_row in debits.iterrows():
                # 借方：简单
                self.results.append({
                    "uid": uid,
                    "src_subject": d_row[subj_col],
                    "contra_subject": target_c,
                    "amount": d_row[self.mapping['debit']],
                    "type": "Nv1"
                })
                # 贷方：被拆分
                self.results.append({
                    "uid": uid,
                    "src_subject": target_c,
                    "contra_subject": d_row[subj_col],
                    "amount": d_row[self.mapping['debit']],
                    "type": "Nv1"
                })

    def _add_to_cluster(self, uid, debits, credits):
        subj_col = self.mapping['subject']
        d_col = self.mapping['debit']
        c_col = self.mapping['credit']

        # === 优化点：模式特征去重且不分借贷 ===
        # 你的要求：直接列示所有去重后的科目，按"、"分开
        all_subjs = set(debits[subj_col].astype(str)) | set(credits[subj_col].astype(str))
        sorted_subjs = sorted(list(all_subjs))
        key_str = "、".join(sorted_subjs)
        
        # 使用 Hash 存储
        key_hash = hashlib.md5(key_str.encode()).hexdigest()
        
        self.complex_clusters[key_hash].append(uid)
        
        # 仅保存第一个作为样本
        if key_hash not in self.cluster_samples:
            # 按科目名聚合金额，这是为了穷举算法的效率
            d_dict = defaultdict(float)
            for _, row in debits.iterrows():
                d_dict[row[subj_col]] += row[d_col]
                
            c_dict = defaultdict(float)
            for _, row in credits.iterrows():
                c_dict[row[subj_col]] += row[c_col]

            self.cluster_samples[key_hash] = {
                "name": key_str,
                "debits": dict(d_dict),
                "credits": dict(c_dict),
                "count": 1,
                "sample_uid": uid
            }
        else:
            self.cluster_samples[key_hash]["count"] += 1