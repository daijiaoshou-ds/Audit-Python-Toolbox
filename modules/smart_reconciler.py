import os
import pandas as pd
import numpy as np
import threading
import time
import re
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
import difflib
from typing import List, Dict, Tuple
import itertools
import warnings
import uuid

# 屏蔽 Pandas 日期警告
warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

# --- 科学计算库 ---
try:
    from scipy.optimize import linear_sum_assignment
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False

# --- 资源路径 ---
from modules.path_manager import get_asset_path

# --- AI 库 ---
try:
    from sentence_transformers import SentenceTransformer, util
    HAS_AI = True
except ImportError:
    HAS_AI = False

# ==================== 核心逻辑引擎 (Backend) ====================

class ReconcilerEngine:
    def __init__(self, log_callback):
        self.log = log_callback
        self.model = None
        self.model_path = get_asset_path(os.path.join("assets", "models", "nlp", "text2vec-base-chinese"))
        
        self.gl_raw_df = None
        self.bank_raw_dfs = {}
        self.gl_columns = []
        self.bank_files_info = [] 

    def load_ai_model(self):
        if not HAS_AI: return
        if self.model: return
        if os.path.exists(self.model_path):
            try:
                # self.log("AI 模型加载中...")
                self.model = SentenceTransformer(self.model_path)
                self.log("✅ AI 模型加载就绪")
            except: pass

    def smart_read_excel(self, path):
        try:
            df_preview = pd.read_excel(path, nrows=30, header=None)
            keywords = ['日期', '交易日', '时间', '摘要', '用途', '户名', '借方', '贷方', '收入', '支出', '金额', '发生额']
            best_idx = 0; max_score = 0
            for idx, row in df_preview.iterrows():
                row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
                score = sum(1 for k in keywords if k in row_str)
                if score > max_score: max_score = score; best_idx = idx
            
            header_row = best_idx if max_score >= 2 else 0
            df = pd.read_excel(path, header=header_row, dtype=object)
            df.dropna(how='all', inplace=True); df.dropna(axis=1, how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            return df, header_row
        except Exception as e: return None, str(e)

    def load_full_gl_data(self, path, cb):
        try:
            cb(0.2, "解析序时账...")
            df, hr = self.smart_read_excel(path)
            if df is None: return False, hr
            df.reset_index(drop=True, inplace=True)
            self.gl_raw_df = df
            self.gl_columns = list(df.columns)
            cb(1.0, f"加载 GL 完成: {len(df)} 行 (表头行: {hr+1})")
            return True, ""
        except Exception as e: return False, str(e)

    def load_bank_file_basic(self, path):
        if path in self.bank_raw_dfs: return list(self.bank_raw_dfs[path].columns)
        df, _ = self.smart_read_excel(path)
        if df is not None:
            self.bank_raw_dfs[path] = df
            return list(df.columns)
        return []

    def remove_bank_file(self, path):
        self.bank_files_info = [x for x in self.bank_files_info if x['path'] != path]
        if path in self.bank_raw_dfs: del self.bank_raw_dfs[path]
        return len(self.bank_files_info)

    def extract_gl_structure(self, l1, l2=None):
        if self.gl_raw_df is None or l1 not in self.gl_raw_df.columns: return []
        return [x for x in self.gl_raw_df[l1].astype(str).str.strip().unique().tolist() if x.lower() != 'nan']

    def filter_gl_details(self, l1, l2, target):
        if self.gl_raw_df is None: return []
        target = str(target).strip()
        mask = self.gl_raw_df[l1].astype(str).str.strip() == target
        filtered = self.gl_raw_df[mask]
        if l2 in filtered.columns:
            return [x for x in filtered[l2].astype(str).str.strip().unique().tolist() if x.lower() != 'nan']
        return []

    def scan_bank_files(self, paths):
        info = []
        existing = [x['path'] for x in self.bank_files_info]
        for p in paths:
            if p in existing: continue
            fname = os.path.basename(p)
            digits = re.findall(r'\d+', fname)
            key = [d for d in digits if len(d)>=3]
            info.append({"path": p, "name": fname, "key_digits": key, "feature_text": fname})
        self.bank_files_info.extend(info)
        return len(self.bank_files_info)

    def auto_match(self, gl_details):
        if not self.bank_files_info: return []
        b_names = [b['name'] for b in self.bank_files_info]
        b_digits = [b['key_digits'] for b in self.bank_files_info]
        b_feats = [b['feature_text'] for b in self.bank_files_info]
        scores = np.zeros((len(gl_details), len(b_names)))

        for i, gl in enumerate(gl_details):
            for j, digs in enumerate(b_digits):
                for d in digs:
                    if d in gl: scores[i][j] += 10.0; break
        
        if self.model:
            try:
                emb_gl = self.model.encode(gl_details, convert_to_tensor=True)
                emb_bank = self.model.encode(b_feats, convert_to_tensor=True)
                scores += util.cos_sim(emb_gl, emb_bank).cpu().numpy()
            except: pass
        else:
            for i, gl in enumerate(gl_details):
                for j, bn in enumerate(b_names): scores[i][j] += difflib.SequenceMatcher(None, gl, bn).ratio()

        res = []
        for i, gl in enumerate(gl_details):
            best = np.argmax(scores[i])
            sc = scores[i][best]
            thresh = 1.0 if np.max(scores) > 5 else 0.4
            res.append((gl, b_names[best], float(sc)) if sc > thresh else (gl, "(未匹配)", 0.0))
        return res

    # ==================== 核心核对算法 (V20.1) ====================

    def _smart_parse_dates(self, series, label):
        s = series.astype(str).str.strip()
        dates_a = pd.to_datetime(s, errors='coerce')
        valid_a = dates_a.notna().sum()
        dates_b = pd.to_datetime(s, dayfirst=True, errors='coerce')
        valid_b = dates_b.notna().sum()
        
        if valid_b > valid_a:
            self.log(f"  ℹ️ [{label}] 识别为 '日/月/年'")
            return dates_b
        
        if valid_a < len(series) * 0.5:
            try:
                numeric_s = pd.to_numeric(series, errors='coerce')
                dates_c = pd.to_datetime(numeric_s, unit='D', origin='1899-12-30', errors='coerce')
                if dates_c.notna().sum() > valid_a:
                    self.log(f"  ℹ️ [{label}] 识别为 Excel 序列号")
                    return dates_c
            except: pass
        return dates_a

    def _normalize_data(self, df, date_col, amt_series, desc_col, voucher_col=None, party_col=None, serial_col=None, label="Data"):
        if amt_series is None: return pd.DataFrame()
        
        df = df.reset_index(drop=True)
        if isinstance(amt_series, pd.Series):
            amt_series = amt_series.reset_index(drop=True)
        
        temp = pd.DataFrame()
        temp['orig_idx'] = df.index
        temp['desc'] = df[desc_col].astype(str).replace('nan', '', regex=False).str.strip() if desc_col and desc_col in df.columns else ""
        temp['party'] = df[party_col].astype(str).replace('nan', '', regex=False).str.strip() if party_col and party_col in df.columns else ""
        temp['voucher'] = df[voucher_col].astype(str).replace('nan', '', regex=False).str.strip() if voucher_col and voucher_col in df.columns else ""
        
        if serial_col and serial_col in df.columns:
            temp['serial'] = df[serial_col].astype(str).replace('nan', '', regex=False).str.strip()
        else:
            temp['serial'] = ""

        temp['date'] = self._smart_parse_dates(df[date_col], label)
        if temp['date'].isna().any(): self.log(f"⚠️ [{label}] 过滤 {temp['date'].isna().sum()} 行无效日期")
        temp = temp.dropna(subset=['date']).copy()
        
        temp['date'] = temp['date'].dt.normalize()
        temp['month'] = temp['date'].dt.to_period('M')
        
        clean_amt = amt_series.astype(str).str.replace(r'[^\d.-]', '', regex=True)
        temp['amount'] = pd.to_numeric(clean_amt, errors='coerce').fillna(0.0).round(2)
        temp = temp[temp['amount'] != 0].copy()
        
        return temp.sort_values('date')

    def find_subset_sum(self, target, pool, limit=6):
        n = len(pool)
        if n > 15: return None 
        pool_vals = [p[1] for p in pool]
        pool_idxs = [p[0] for p in pool]
        for r in range(2, min(n, limit) + 1):
            for combo_idx in itertools.combinations(range(n), r):
                if abs(sum([pool_vals[i] for i in combo_idx]) - target) < 0.02: 
                    return [pool_idxs[i] for i in combo_idx]
        return None

    def _is_name_match(self, name_a, name_b, threshold):
        if not name_a or not name_b: return False
        na = name_a.strip(); nb = name_b.strip()
        if not na or not nb: return False
        if na in nb or nb in na: return True
        return difflib.SequenceMatcher(None, na, nb).ratio() >= threshold

    def execute_reconciliation(self, mapping_dict, gl_cfg, bank_cfgs, output_path, strategy_cfg, stop_event=None):
        writer = pd.ExcelWriter(output_path, engine='openpyxl')
        summary_data = []
        name_threshold = strategy_cfg.get('name_threshold', 0.3)
        def get_gid(): return uuid.uuid4().hex[:8]

        for gl_sub_name, bank_path_key in mapping_dict.items():
            if stop_event and stop_event.is_set():
                self.log(">>> 用户强制停止任务！")
                return False, "任务已终止"

            self.log(f"--- 核对: {gl_sub_name} ---")
            
            # 1. GL
            s_l1 = self.gl_raw_df[gl_cfg['l1']].astype(str).str.strip()
            s_l2 = self.gl_raw_df[gl_cfg['l2']].astype(str).str.strip()
            target_l1 = str(gl_cfg['target']).strip(); target_l2 = str(gl_sub_name).strip()
            gl_mask = (s_l1 == target_l1) & (s_l2 == target_l2)
            gl_subset = self.gl_raw_df[gl_mask].copy()
            if gl_subset.empty: continue

            gl_val = pd.to_numeric(gl_subset[gl_cfg['debit']], errors='coerce').fillna(0) - \
                     pd.to_numeric(gl_subset[gl_cfg['credit']], errors='coerce').fillna(0)
            df_gl = self._normalize_data(gl_subset, gl_cfg['date'], gl_val, gl_cfg['desc'], voucher_col=gl_cfg['voucher'], party_col=gl_cfg['party'], serial_col=None, label="GL")

            # 2. Bank
            fname = os.path.basename(bank_path_key)
            if fname not in bank_cfgs: continue
            b_cfg = bank_cfgs[fname]
            df_bank_raw = self.bank_raw_dfs.get(bank_path_key)
            if df_bank_raw is None: continue

            if b_cfg['mode'] == '2col':
                b_val = pd.to_numeric(df_bank_raw[b_cfg['credit']], errors='coerce').fillna(0) - \
                        pd.to_numeric(df_bank_raw[b_cfg['debit']], errors='coerce').fillna(0)
            else:
                b_val = pd.to_numeric(df_bank_raw[b_cfg['credit']], errors='coerce').fillna(0)
            
            df_bank = self._normalize_data(df_bank_raw, b_cfg['date'], b_val, b_cfg['desc'], party_col=b_cfg['party'], serial_col=b_cfg['serial'], label="Bank")

            self.log(f"  GL: {len(df_gl)} 笔 | Bank: {len(df_bank)} 笔")

            # 3. Matching
            matches = [] 
            pool_gl = set(df_gl.index)
            pool_bank = set(df_bank.index)

            # >>> P1: 精确 (Exact)
            if stop_event and stop_event.is_set(): return False, "任务已终止"
            map_gl = {}
            for idx in pool_gl:
                row = df_gl.loc[idx]
                map_gl.setdefault((row['amount'], row['date']), []).append(idx)
            cnt_p1 = 0
            for idx_b in list(pool_bank):
                r = df_bank.loc[idx_b]
                key = (r['amount'], r['date'])
                if key in map_gl and map_gl[key]:
                    idx_g = map_gl[key].pop(0)
                    matches.append((idx_g, idx_b, "P1-精确", get_gid()))
                    pool_gl.remove(idx_g); pool_bank.remove(idx_b); cnt_p1 += 1
            self.log(f"  > P1 精确: {cnt_p1}")

            # >>> P2: 邻近 (Proximity)
            if stop_event and stop_event.is_set(): return False, "任务已终止"
            map_gl_amt = {}
            for idx in pool_gl: map_gl_amt.setdefault(df_gl.loc[idx, 'amount'], []).append(idx)
            cnt_p2 = 0
            for idx_b in list(pool_bank):
                row = df_bank.loc[idx_b]
                amt = row['amount']; date_b = row['date']
                if amt in map_gl_amt and map_gl_amt[amt]:
                    candidates = map_gl_amt[amt]
                    best_i = None; min_diff = 3
                    for i, idx_g in enumerate(candidates):
                        date_g = df_gl.loc[idx_g, 'date']
                        if pd.isna(date_g): continue
                        if date_g.month != date_b.month or date_g.year != date_b.year: continue
                        diff = abs((date_g - date_b).days)
                        if diff <= 2 and diff < min_diff: min_diff = diff; best_i = i
                    if best_i is not None:
                        idx_g = candidates.pop(best_i)
                        matches.append((idx_g, idx_b, f"P2-邻近{min_diff}天", get_gid()))
                        pool_gl.remove(idx_g); pool_bank.remove(idx_b); cnt_p2 += 1
            self.log(f"  > P2 邻近: {cnt_p2}")

            # >>> P3: 同月 (Same Month)
            if stop_event and stop_event.is_set(): return False, "任务已终止"
            cnt_p3 = 0
            months = set(df_gl.loc[list(pool_gl), 'month'].unique()) | set(df_bank.loc[list(pool_bank), 'month'].unique())
            for m in months:
                gl_m = [i for i in pool_gl if df_gl.loc[i, 'month'] == m]
                bk_m = [i for i in pool_bank if df_bank.loc[i, 'month'] == m]
                gl_amt = {}; bk_amt = {}
                for i in gl_m: gl_amt.setdefault(df_gl.loc[i, 'amount'], []).append(i)
                for i in bk_m: bk_amt.setdefault(df_bank.loc[i, 'amount'], []).append(i)
                for amt in set(gl_amt) & set(bk_amt):
                    gs, bs = gl_amt[amt], bk_amt[amt]
                    gs.sort(key=lambda x: df_gl.loc[x, 'date'])
                    bs.sort(key=lambda x: df_bank.loc[x, 'date'])
                    count = min(len(gs), len(bs))
                    for k in range(count):
                        matches.append((gs[k], bs[k], "P3-同月跨期", get_gid()))
                        pool_gl.remove(gs[k]); pool_bank.remove(bs[k]); cnt_p3 += 1
            self.log(f"  > P3 同月: {cnt_p3}")

            # >>> P4: 聚合 (Aggregation 4-Layers) <<<
            cnt_p4 = 0
            if strategy_cfg.get('aggregation'):
                sub_strats = ['day_homo', 'month_homo', 'day_all', 'month_all']
                
                for strat in sub_strats:
                    if stop_event and stop_event.is_set(): return False, "任务已终止"
                    bank_groups = {}
                    for idx in pool_bank:
                        r = df_bank.loc[idx]
                        if 'homo' in strat:
                            desc_key = str(r['desc'])[:10] if pd.notna(r['desc']) else "UNK"
                            party_key = str(r['party']) if pd.notna(r['party']) else "UNK"
                        else:
                            desc_key = "ALL"; party_key = "ALL"
                        
                        if 'day' in strat:
                            time_key = r['date']
                        else:
                            time_key = r['month']
                            
                        key = (time_key, party_key, desc_key)
                        bank_groups.setdefault(key, []).append(idx)
                    
                    for key, b_indices in list(bank_groups.items()):
                        if len(b_indices) < 2: continue
                        
                        grp_amt = round(sum(df_bank.loc[i, 'amount'] for i in b_indices), 2)
                        grp_time = key[0] # date or month
                        
                        best_g = None
                        for g_idx in pool_gl:
                            g_row = df_gl.loc[g_idx]
                            if abs(g_row['amount'] - grp_amt) < 0.01:
                                match_time = False
                                if 'day' in strat:
                                    # Day Aggregation -> Match GL by Same Month (宽松匹配，解决 15日凑 25日)
                                    # 只要 GL 的月份 == 聚合 key 的月份即可
                                    if g_row['month'] == grp_time.to_period('M'): match_time = True
                                else:
                                    # Month Aggregation -> Match GL by Same Month
                                    if g_row['month'] == grp_time: match_time = True
                                
                                if match_time:
                                    best_g = g_idx; break
                        
                        if best_g:
                            gid = get_gid()
                            label = f"P4-聚合({strat.split('_')[0]})"
                            matches.append((best_g, b_indices[0], f"{label}-主", gid))
                            for k in range(1, len(b_indices)): matches.append((None, b_indices[k], f"{label}-子", gid))
                            pool_gl.remove(best_g)
                            for bi in b_indices: pool_bank.remove(bi)
                            cnt_p4 += 1
            self.log(f"  > P4 聚合: {cnt_p4}")

            # >>> P5: 暴力凑数 (Subset) <<<
            cnt_p5 = 0
            if strategy_cfg.get('subset'):
                months = set(df_gl.loc[list(pool_gl), 'month'].unique())
                for m in months:
                    if stop_event and stop_event.is_set(): return False, "任务已终止"
                    gl_m_raw = [i for i in pool_gl if df_gl.loc[i, 'month'] == m]
                    bk_m_raw = [i for i in pool_bank if df_bank.loc[i, 'month'] == m]
                    if not gl_m_raw or not bk_m_raw: continue

                    # A. 引导凑数
                    for bi in list(bk_m_raw):
                        if bi not in pool_bank: continue
                        tgt = df_bank.loc[bi, 'amount']; bk_party = str(df_bank.loc[bi, 'party'])
                        candidates = []
                        for gi in gl_m_raw:
                            if gi not in pool_gl: continue
                            g_desc = str(df_gl.loc[gi, 'desc']); g_party = str(df_gl.loc[gi, 'party'])
                            if (bk_party and len(bk_party)>1 and bk_party in g_desc) or self._is_name_match(g_party, bk_party, name_threshold):
                                candidates.append(gi)
                        
                        if len(candidates) >= 2:
                            pool_tuples = [(i, df_gl.loc[i, 'amount']) for i in candidates]
                            res = self.find_subset_sum(tgt, pool_tuples)
                            if res:
                                gid = get_gid()
                                matches.append((res[0], bi, "P5-引导凑数(N:1)", gid))
                                for k in range(1, len(res)): matches.append((res[k], None, "P5-引导凑数(子)", gid))
                                pool_bank.remove(bi)
                                for gi in res: pool_gl.remove(gi)
                                cnt_p5 += 1; continue

                    # B. 暴力凑数
                    for scope in [0, 1, 2]: 
                        # Dir A: 1 GL vs N Bank
                        for gi in list(gl_m_raw):
                            if gi not in pool_gl: continue
                            tgt = df_gl.loc[gi, 'amount']; g_date = df_gl.loc[gi, 'date']
                            if scope==0: bk_pool = [i for i in bk_m_raw if i in pool_bank and df_bank.loc[i,'date']==g_date]
                            elif scope==1: bk_pool = [i for i in bk_m_raw if i in pool_bank and abs((df_bank.loc[i,'date']-g_date).days)<=7]
                            else: bk_pool = [i for i in bk_m_raw if i in pool_bank]
                            
                            # === 【修复点】 === 确保 pool_tuples 初始化
                            pool_tuples = [(i, df_bank.loc[i, 'amount']) for i in bk_pool]
                            res = self.find_subset_sum(tgt, pool_tuples)
                            if res:
                                gid = get_gid()
                                matches.append((gi, res[0], f"P5-凑数(1:N)-S{scope}", gid))
                                for k in range(1, len(res)): matches.append((None, res[k], "P5-凑数(子)", gid))
                                pool_gl.remove(gi); 
                                for bi in res: pool_bank.remove(bi)
                                cnt_p5 += 1

                        # Dir B: N GL vs 1 Bank
                        for bi in list(bk_m_raw):
                            if bi not in pool_bank: continue
                            tgt = df_bank.loc[bi, 'amount']; b_date = df_bank.loc[bi, 'date']
                            if scope==0: gl_pool = [i for i in gl_m_raw if i in pool_gl and df_gl.loc[i,'date']==b_date]
                            elif scope==1: gl_pool = [i for i in gl_m_raw if i in pool_gl and abs((df_gl.loc[i,'date']-b_date).days)<=7]
                            else: gl_pool = [i for i in gl_m_raw if i in pool_gl]
                            
                            # === 【修复点】 === 确保 pool_tuples 初始化
                            pool_tuples = [(i, df_gl.loc[i, 'amount']) for i in gl_pool]
                            res = self.find_subset_sum(tgt, pool_tuples)
                            if res:
                                gid = get_gid()
                                matches.append((res[0], bi, f"P5-凑数(N:1)-S{scope}", gid))
                                for k in range(1, len(res)): matches.append((res[k], None, "P5-凑数(子)", gid))
                                pool_bank.remove(bi); 
                                for gi in res: pool_gl.remove(gi)
                                cnt_p5 += 1

            self.log(f"  > P5 智能凑数: {cnt_p5}")

            # --- 输出 ---
            res_rows = []
            for g, b, t, gid in matches:
                if g is not None:
                    gr = df_gl.loc[g]
                    g_d, g_v, g_a, g_desc, g_pty = gr['date'], gr.get('voucher',''), gr['amount'], gr['desc'], gr['party']
                else:
                    g_d, g_v, g_a, g_desc, g_pty = None, None, None, "(聚合/凑数子项)", None
                
                if b is not None:
                    br = df_bank.loc[b]
                    b_d, b_a, b_desc, b_pty, b_ser = br['date'], br['amount'], br['desc'], br['party'], br['serial']
                else:
                    b_d, b_a, b_desc, b_pty, b_ser = None, None, None, None, None

                res_rows.append({
                    "匹配组ID": gid,
                    "GL_日期": g_d, "GL_凭证": g_v, "GL_金额": g_a, "GL_摘要": g_desc, "GL_客商": g_pty,
                    "Bank_日期": b_d, "Bank_金额": b_a, "Bank_摘要": b_desc, "Bank_交易方": b_pty, "Bank_流水号": b_ser,
                    "匹配类型": t, "差异": 0
                })
            
            for g in pool_gl:
                gr = df_gl.loc[g]
                res_rows.append({"匹配组ID": "未达_GL", "GL_日期": gr['date'], "GL_凭证": gr.get('voucher',''), "GL_金额": gr['amount'], "GL_摘要": gr['desc'], "GL_客商": gr['party'], "匹配类型": "企业已记银行未记", "差异": gr['amount']})
            for b in pool_bank:
                br = df_bank.loc[b]
                res_rows.append({"匹配组ID": "未达_BK", "Bank_日期": br['date'], "Bank_金额": br['amount'], "Bank_摘要": br['desc'], "Bank_交易方": br['party'], "Bank_流水号": br['serial'], "匹配类型": "银行已记企业未记", "差异": -br['amount']})

            safe_name = re.sub(r'[\\/*?:\[\]]', '', str(gl_sub_name))[:30]
            pd.DataFrame(res_rows).to_excel(writer, sheet_name=safe_name, index=False)
            
            summary_data.append({"科目": gl_sub_name, "匹配数": len(matches), "未达GL": len(pool_gl), "未达Bank": len(pool_bank)})

        if summary_data: pd.DataFrame(summary_data).to_excel(writer, sheet_name="核对汇总", index=False)
        writer.close()
        return True, f"完成! 结果: {output_path}"

# ==================== 界面模块 (Frontend) ====================

class SmartReconcilerModule:
    def __init__(self):
        self.name = "银行流水核查"
        self.engine = None
        self.gl_path = ""; self.bank_paths = []
        self.bank_ui_rows = {} 

    def render(self, parent_frame):
        if not self.engine:
            self.engine = ReconcilerEngine(self.log)
            threading.Thread(target=self.engine.load_ai_model, daemon=True).start()

        for w in parent_frame.winfo_children(): w.destroy()

        self.main_scroll = ctk.CTkScrollableFrame(parent_frame, fg_color="#F2F4F8", scrollbar_button_color="#D0D0D0")
        self.main_scroll.pack(fill="both", expand=True)

        ctk.CTkLabel(self.main_scroll, text="银行流水核查", font=("Microsoft YaHei", 24, "bold"), text_color="#333").pack(anchor="w", padx=20, pady=(20, 10))

        self.create_gl_section(self.main_scroll)
        self.create_bank_section(self.main_scroll)
        self.create_bank_config_section(self.main_scroll) 
        self.create_mapping_section(self.main_scroll)
        self.create_match_section(self.main_scroll)

        ctk.CTkLabel(self.main_scroll, text="执行日志", font=("Arial", 12, "bold"), text_color="#555").pack(anchor="w", padx=25, pady=(10,5))
        self.log_box = ctk.CTkTextbox(self.main_scroll, height=250, fg_color="white", text_color="#333", border_color="#CCC", border_width=1)
        self.log_box.pack(fill="x", padx=20, pady=(0, 30))

    def _frame(self, parent):
        f = ctk.CTkFrame(parent, fg_color="white", corner_radius=8, border_width=2, border_color="#E0E0E0")
        f.pack(fill="x", padx=20, pady=10)
        return f

    def create_gl_section(self, parent):
        f = self._frame(parent)
        ctk.CTkLabel(f, text="1. 序时账配置", font=("Microsoft YaHei", 15, "bold"), text_color="#007AFF").pack(anchor="w", padx=15, pady=15)
        
        row1 = ctk.CTkFrame(f, fg_color="transparent"); row1.pack(fill="x", padx=15)
        self.btn_load_gl = ctk.CTkButton(row1, text="导入序时账...", command=self.load_gl_file, width=120, fg_color="#F0F5FF", text_color="#007AFF", border_width=1, border_color="#007AFF")
        self.btn_load_gl.pack(side="left")
        self.lbl_gl_info = ctk.CTkLabel(row1, text="未选择", text_color="#999"); self.lbl_gl_info.pack(side="left", padx=10)
        
        self.progress_gl = ctk.CTkProgressBar(f, height=6); 
        
        col_frame = ctk.CTkFrame(f, fg_color="#FAFAFA", corner_radius=6); col_frame.pack(fill="x", padx=15, pady=15)
        grid = ctk.CTkFrame(col_frame, fg_color="transparent"); grid.pack(fill="x", padx=10, pady=10)
        grid.grid_columnconfigure((1,3,5), weight=1)

        def cb(r, c, txt):
            ctk.CTkLabel(grid, text=txt, text_color="#333").grid(row=r, column=c*2, padx=5, sticky="e")
            b = ctk.CTkComboBox(grid, width=130, fg_color="white", button_color="#DDD", text_color="#333", dropdown_fg_color="white", dropdown_text_color="#333")
            b.set(""); b.grid(row=r, column=c*2+1, padx=5, sticky="ew"); return b

        self.combo_l1 = cb(0,0,"一级科目:"); self.combo_l1.configure(command=self.on_l1_col_change)
        self.combo_l2 = cb(0,1,"明细科目:")
        
        ctk.CTkLabel(grid, text="筛选目标:", text_color="#d63031", font=("Arial", 12, "bold")).grid(row=0, column=4, padx=5, sticky="e")
        self.combo_target_l1 = ctk.CTkComboBox(grid, width=150, command=self.on_target_subject_change, fg_color="white", button_color="#DDD", text_color="#333", dropdown_fg_color="white", dropdown_text_color="#333")
        self.combo_target_l1.set(""); self.combo_target_l1.grid(row=0, column=5, padx=5, sticky="ew")

        self.combo_date = cb(1,0,"日期列:")
        self.combo_debit = cb(1,1,"借方金额:")
        self.combo_credit = cb(1,2,"贷方金额:")
        self.combo_desc = cb(2,0,"摘要列:")
        self.combo_voucher = cb(2,1,"凭证号:")
        self.combo_party = cb(2,2,"客商名称(选填):")

    def create_bank_section(self, parent):
        f = self._frame(parent)
        ctk.CTkLabel(f, text="2. 银行流水导入", font=("Microsoft YaHei", 15, "bold"), text_color="#00b894").pack(anchor="w", padx=15, pady=15)
        row = ctk.CTkFrame(f, fg_color="transparent"); row.pack(fill="x", padx=15, pady=(0,15))
        ctk.CTkButton(row, text="添加文件", command=self.add_bank_files, width=100, fg_color="white", text_color="#333", border_width=1, border_color="#CCC").pack(side="left")
        ctk.CTkButton(row, text="添加文件夹", command=self.add_bank_dir, width=100, fg_color="white", text_color="#333", border_width=1, border_color="#CCC").pack(side="left", padx=10)
        ctk.CTkButton(row, text="清空所有", command=self.clear_bank_files, width=80, fg_color="white", text_color="#d63031", border_width=1, border_color="#CCC").pack(side="right", padx=10)
        self.lbl_bank_count = ctk.CTkLabel(row, text="0 文件", text_color="#666"); self.lbl_bank_count.pack(side="left", padx=5)

    def create_bank_config_section(self, parent):
        f = self._frame(parent)
        head = ctk.CTkFrame(f, fg_color="transparent"); head.pack(fill="x", padx=15, pady=15)
        ctk.CTkLabel(head, text="3. 银行流水字段配置", font=("Microsoft YaHei", 15, "bold"), text_color="#e17055").pack(side="left")
        ctk.CTkButton(head, text="刷新配置表", command=self.refresh_bank_config_ui, width=100, height=28, fg_color="#F0F0F0", text_color="#333").pack(side="right")
        
        container = ctk.CTkFrame(f, fg_color="#FFF8F0", height=250)
        container.pack(fill="x", padx=15, pady=(0,15)); container.pack_propagate(False)
        v_scroll = ctk.CTkScrollbar(container, orientation="vertical", fg_color="#FFF8F0", button_color="#D0D0D0", button_hover_color="#C0C0C0")
        h_scroll = ctk.CTkScrollbar(container, orientation="horizontal", fg_color="#FFF8F0", button_color="#D0D0D0", button_hover_color="#C0C0C0")
        v_scroll.pack(side="right", fill="y"); h_scroll.pack(side="bottom", fill="x")
        self.canvas = tk.Canvas(container, bg="#FFF8F0", bd=0, highlightthickness=0, yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        v_scroll.configure(command=self.canvas.yview); h_scroll.configure(command=self.canvas.xview)
        self.inner_frame = ctk.CTkFrame(self.canvas, fg_color="#FFF8F0")
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def refresh_bank_config_ui(self):
        for w in self.inner_frame.winfo_children(): w.destroy()
        self.bank_ui_rows = {}
        if not self.engine.bank_files_info:
            ctk.CTkLabel(self.inner_frame, text="无文件", text_color="#999").pack(pady=20, padx=20); return

        h_frame = ctk.CTkFrame(self.inner_frame, fg_color="transparent"); h_frame.pack(fill="x", pady=5, padx=5)
        headers = ["文件名", "模式", "日期", "贷(收)", "借(支)", "摘要", "交易方", "流水号(新增)"]
        widths = [180, 100, 100, 100, 100, 100, 120, 120] 
        for i, (h, w) in enumerate(zip(headers, widths)):
            f = ctk.CTkFrame(h_frame, fg_color="transparent", width=w, height=30)
            f.pack_propagate(False); f.pack(side="left", padx=2)
            ctk.CTkLabel(f, text=h, font=("Arial", 11, "bold"), text_color="#555").pack(expand=True)

        for info in self.engine.bank_files_info:
            path = info['path']; fname = info['name']
            cols = [""] + self.engine.load_bank_file_basic(path)
            row = ctk.CTkFrame(self.inner_frame, fg_color="white", corner_radius=6); row.pack(fill="x", pady=2, padx=5)

            f1 = ctk.CTkFrame(row, fg_color="transparent", width=widths[0], height=40)
            f1.pack_propagate(False); f1.pack(side="left", padx=2)
            ctk.CTkLabel(f1, text=fname, font=("Arial", 10), anchor="w").pack(side="left", padx=5)
            ctk.CTkButton(f1, text="×", width=20, height=20, fg_color="#FFF0F0", text_color="red", command=lambda p=path: self.delete_single_bank(p)).pack(side="right")

            f2 = ctk.CTkFrame(row, fg_color="transparent", width=widths[1], height=40)
            f2.pack_propagate(False); f2.pack(side="left", padx=2)
            mode_var = ctk.StringVar(value="2col")
            cb_mode = ctk.CTkComboBox(f2, values=["双列(收/支)", "单列(正/负)"], width=110, variable=mode_var, fg_color="white", text_color="#333", dropdown_fg_color="white")
            cb_mode.pack(expand=True)

            def mk_cb(parent_w):
                f = ctk.CTkFrame(row, fg_color="transparent", width=parent_w, height=40)
                f.pack_propagate(False); f.pack(side="left", padx=2)
                cb = ctk.CTkComboBox(f, values=cols, width=parent_w-10, fg_color="white", text_color="#333", dropdown_fg_color="white")
                cb.set(""); cb.pack(expand=True); return cb

            cb_dt = mk_cb(widths[2]); cb_cr = mk_cb(widths[3]); cb_dr = mk_cb(widths[4]); cb_desc = mk_cb(widths[5]); cb_party = mk_cb(widths[6]); cb_serial = mk_cb(widths[7])

            def on_mode_change(val, dr=cb_dr, cr=cb_cr):
                if val == "单列(正/负)": dr.configure(state="disabled", fg_color="#EEE"); cr.set("")
                else: dr.configure(state="normal", fg_color="white")
            cb_mode.configure(command=on_mode_change)

            for c in cols:
                cs = str(c)
                if not cs: continue
                if "日期" in cs or "交易日" in cs: cb_dt.set(c)
                if "贷" in cs or "收" in cs: cb_cr.set(c)
                if "借" in cs or "支" in cs: cb_dr.set(c)
                if "摘要" in cs or "用途" in cs: cb_desc.set(c)
                if "对方" in cs or "户名" in cs or "名称" in cs: cb_party.set(c)
                if "流水" in cs or "凭证" in cs or "单号" in cs or "票据" in cs: cb_serial.set(c)

            self.bank_ui_rows[fname] = {"mode": mode_var, "date": cb_dt, "credit": cb_cr, "debit": cb_dr, "desc": cb_desc, "party": cb_party, "serial": cb_serial}

    def delete_single_bank(self, path):
        n = self.engine.remove_bank_file(path)
        self.lbl_bank_count.configure(text=f"{n} 文件")
        self.refresh_bank_config_ui()

    def clear_bank_files(self):
        self.engine.bank_files_info = []
        self.engine.bank_raw_dfs = {}
        self.lbl_bank_count.configure(text="0 文件")
        self.refresh_bank_config_ui()

    def create_mapping_section(self, parent):
            f = self._frame(parent)
            head = ctk.CTkFrame(f, fg_color="transparent"); head.pack(fill="x", padx=15, pady=15)
            ctk.CTkLabel(head, text="4. 建立映射关系", font=("Microsoft YaHei", 15, "bold"), text_color="#6c5ce7").pack(side="left")
            
            btn_box = ctk.CTkFrame(head, fg_color="transparent"); btn_box.pack(side="right")
            ctk.CTkButton(btn_box, text="🤖 AI 自动匹配", command=self.run_ai_mapping, width=100, height=30, fg_color="#6c5ce7").pack(side="left", padx=5)
            ctk.CTkButton(btn_box, text="导出", command=self.export_mapping, width=60, height=30, fg_color="#F0F0F0", text_color="#333").pack(side="left", padx=5)
            ctk.CTkButton(btn_box, text="导入", command=self.import_mapping, width=60, height=30, fg_color="#F0F0F0", text_color="#333").pack(side="left", padx=5)
            
            # === 【UI 修复】 双向滚动容器 ===
            container = ctk.CTkFrame(f, fg_color="#FAFAFA", height=200)
            container.pack(fill="x", padx=15, pady=(0,15))
            container.pack_propagate(False) # 固定高度
            
            # 滚动条
            v_scroll = ctk.CTkScrollbar(container, orientation="vertical", fg_color="#FAFAFA", button_color="#D0D0D0", button_hover_color="#C0C0C0")
            h_scroll = ctk.CTkScrollbar(container, orientation="horizontal", fg_color="#FAFAFA", button_color="#D0D0D0", button_hover_color="#C0C0C0")
            v_scroll.pack(side="right", fill="y")
            h_scroll.pack(side="bottom", fill="x")
            
            # 画布
            self.map_canvas = tk.Canvas(container, bg="#FAFAFA", bd=0, highlightthickness=0, 
                                        yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
            self.map_canvas.pack(side="left", fill="both", expand=True)
            
            v_scroll.configure(command=self.map_canvas.yview)
            h_scroll.configure(command=self.map_canvas.xview)
            
            # 内部 Frame (存放表格内容)
            self.frame_map_table = ctk.CTkFrame(self.map_canvas, fg_color="#FAFAFA")
            self.map_canvas_window = self.map_canvas.create_window((0, 0), window=self.frame_map_table, anchor="nw")
            
            # 动态更新滚动区域
            def on_frame_config(event):
                self.map_canvas.configure(scrollregion=self.map_canvas.bbox("all"))
            self.frame_map_table.bind("<Configure>", on_frame_config)

    def create_match_section(self, parent):
        f = self._frame(parent)
        ctk.CTkLabel(f, text="5. 核心核对", font=("Microsoft YaHei", 15, "bold"), text_color="#d63031").pack(anchor="w", padx=15, pady=15)
        opt = ctk.CTkFrame(f, fg_color="transparent"); opt.pack(fill="x", padx=15, pady=(0,15))
        self.var_exact = ctk.BooleanVar(value=True)
        self.var_hungarian = ctk.BooleanVar(value=True)
        self.var_subset = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opt, text="基础匹配 (精确+邻近+同月)", state="disabled", text_color="#333").select(); 
        ctk.CTkCheckBox(opt, text="高级: 同质聚合 (拆单汇总)", variable=self.var_hungarian, text_color="#333").pack(side="left", padx=10)
        ctk.CTkCheckBox(opt, text="高级: 智能凑数 (暴力计算)", variable=self.var_subset, text_color="#d63031").pack(side="left", padx=10)
        
        row_slider = ctk.CTkFrame(f, fg_color="transparent"); row_slider.pack(fill="x", padx=15, pady=(5,15))
        ctk.CTkLabel(row_slider, text="名称相似度阈值:", text_color="#666").pack(side="left", padx=(0,10))
        self.lbl_thresh = ctk.CTkLabel(row_slider, text="0.3", text_color="#007AFF"); self.lbl_thresh.pack(side="right", padx=10)
        self.slider_thresh = ctk.CTkSlider(row_slider, from_=0.1, to=1.0, number_of_steps=9, command=lambda v: self.lbl_thresh.configure(text=f"{v:.1f}"))
        self.slider_thresh.set(0.3); self.slider_thresh.pack(fill="x", padx=10)

        # === 【修改点 3】 申请中断信号 ===
        self.btn_start_match = ctk.CTkButton(f, text="🚀 生成审计底稿", width=200, height=45, font=("Microsoft YaHei", 16, "bold"), fg_color="#d63031", state="disabled", command=self.start_core_matching)
        self.btn_start_match.pack(pady=(0, 20))

    # --- 交互 ---
    def log(self, msg):
        try: self.log_box.insert("end", f"> {msg}\n"); self.log_box.see("end")
        except: pass

    def load_gl_file(self):
        p = filedialog.askopenfilename()
        if not p: return
        self.gl_path = p; self.lbl_gl_info.configure(text=os.path.basename(p))
        self.progress_gl.pack(fill="x", padx=15, pady=(0,10)); self.progress_gl.set(0); self.btn_load_gl.configure(state="disabled")
        def t():
            def cb(val, txt): self.progress_gl.set(val); self.log(txt)
            ok, msg = self.engine.load_full_gl_data(p, cb)
            self.btn_load_gl.configure(state="normal"); self.progress_gl.pack_forget()
            if ok:
                cols = [""] + self.engine.gl_columns
                for b in [self.combo_l1, self.combo_l2, self.combo_date, self.combo_debit, self.combo_credit, self.combo_desc, self.combo_voucher, self.combo_party]: b.configure(values=cols)
                for c in cols:
                    if not c: continue
                    if "一级" in c: self.combo_l1.set(c)
                    if "明细" in c: self.combo_l2.set(c)
                    if "日期" in c: self.combo_date.set(c)
                    if "借方" in c: self.combo_debit.set(c)
                    if "贷方" in c: self.combo_credit.set(c)
                    if "摘要" in c: self.combo_desc.set(c)
                    if "凭证" in c: self.combo_voucher.set(c)
                    if "客商" in c or "客户" in c or "供应商" in c: self.combo_party.set(c)
                if self.combo_l1.get(): self.on_l1_col_change(self.combo_l1.get())
            else: messagebox.showerror("错误", msg)
        threading.Thread(target=t, daemon=True).start()

    def on_l1_col_change(self, col):
        items = self.engine.extract_gl_structure(col, None)
        self.combo_target_l1.configure(values=items)
        for i in items:
            if "银行" in str(i): self.combo_target_l1.set(i); break

    def on_target_subject_change(self, val): self.log(f"目标: {val}")
    
    def add_bank_files(self):
        ps = filedialog.askopenfilenames()
        if ps:
            n = self.engine.scan_bank_files(ps); self.lbl_bank_count.configure(text=f"{n} 文件")
            self.log(f"添加 {len(ps)} 个文件"); self.refresh_bank_config_ui()

    def add_bank_dir(self):
        d = filedialog.askdirectory()
        if d:
            fs = [os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith(('.xls','.xlsx')) and not f.startswith("~$")]
            n = self.engine.scan_bank_files(fs); self.lbl_bank_count.configure(text=f"{n} 文件"); self.refresh_bank_config_ui()

    def run_ai_mapping(self):
        l1, l2, tg = self.combo_l1.get(), self.combo_l2.get(), self.combo_target_l1.get()
        if not l1 or not l2 or not tg: return messagebox.showwarning("提示", "请配置序时账")
        if not self.bank_ui_rows: return messagebox.showwarning("提示", "请配置银行流水")
        self.log("AI 匹配中...")
        gl_dtl = self.engine.filter_gl_details(l1, l2, tg)
        def t():
            res = self.engine.auto_match(gl_dtl); self.render_mapping_ui(res); self.btn_start_match.configure(state="normal")
        threading.Thread(target=t, daemon=True).start()

    def render_mapping_ui(self, res):
            for w in self.frame_map_table.winfo_children(): w.destroy()
            
            # Grid 配置
            # 这里不设置 weight，而是设置固定的 minsize，强行撑开宽度以触发横向滚动
            
            self.mapping_combos = []
            b_opts = ["(未匹配)"] + list(self.bank_ui_rows.keys())
            
            # 表头 (设置较宽的 minsize)
            headers = ["GL 明细科目", "Bank 流水文件 (下拉选择)", "AI 置信度"]
            widths = [250, 300, 100] # 给足宽度
            
            for i, (h, w) in enumerate(zip(headers, widths)):
                f = ctk.CTkFrame(self.frame_map_table, fg_color="transparent", width=w, height=30)
                f.pack_propagate(False)
                f.grid(row=0, column=i, padx=5, pady=5)
                ctk.CTkLabel(f, text=h, font=("Arial", 11, "bold"), text_color="#555").pack(side="left")

            for i, (gl, bk, sc) in enumerate(res):
                r = i + 1
                # 1. GL 科目
                f1 = ctk.CTkFrame(self.frame_map_table, fg_color="white", width=widths[0], height=35)
                f1.pack_propagate(False)
                f1.grid(row=r, column=0, padx=5, pady=2)
                ctk.CTkLabel(f1, text=gl, anchor="w").pack(side="left", padx=5, fill="x")
                
                # 2. 下拉框
                f2 = ctk.CTkFrame(self.frame_map_table, fg_color="transparent", width=widths[1], height=35)
                f2.pack_propagate(False)
                f2.grid(row=r, column=1, padx=5, pady=2)
                
                cb = ctk.CTkComboBox(f2, values=b_opts, width=280, fg_color="white", text_color="#333", dropdown_fg_color="white", dropdown_text_color="#333")
                cb.set(bk if bk in b_opts else "(未匹配)")
                cb.pack(side="left")
                self.mapping_combos.append((gl, cb))
                
                # 3. 分数
                f3 = ctk.CTkFrame(self.frame_map_table, fg_color="transparent", width=widths[2], height=35)
                f3.pack_propagate(False)
                f3.grid(row=r, column=2, padx=5, pady=2)
                
                col = "green" if sc > 5 else "orange"
                ctk.CTkLabel(f3, text="自动" if sc>5 else f"{sc:.2f}", text_color=col).pack(side="left", padx=5)

    def export_mapping(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            data = [{"GL科目": gl, "Bank文件": c.get()} for gl, c in self.mapping_combos]
            pd.DataFrame(data).to_excel(p, index=False); messagebox.showinfo("成功", "已导出")

    def import_mapping(self):
        p = filedialog.askopenfilename()
        if p:
            try:
                df = pd.read_excel(p)
                res = [(str(r["GL科目"]), str(r["Bank文件"]), 10.0) for _, r in df.iterrows()]
                self.render_mapping_ui(res); self.btn_start_match.configure(state="normal")
            except: pass

    def start_core_matching(self):
        gl_cfg = {
            'l1': self.combo_l1.get(), 'l2': self.combo_l2.get(), 'target': self.combo_target_l1.get(),
            'date': self.combo_date.get(), 'debit': self.combo_debit.get(), 'credit': self.combo_credit.get(), 
            'desc': self.combo_desc.get(), 'voucher': self.combo_voucher.get(), 'party': self.combo_party.get()
        }
        bank_cfgs = {}
        for fname, widgets in self.bank_ui_rows.items():
            mode = "1col" if widgets['mode'].get() == "单列(正/负)" else "2col"
            bank_cfgs[fname] = {
                'mode': mode, 'date': widgets['date'].get(), 'credit': widgets['credit'].get(), 
                'debit': widgets['debit'].get(), 'desc': widgets['desc'].get(), 
                'party': widgets['party'].get(),
                'serial': widgets['serial'].get()
            }
            if not bank_cfgs[fname]['date'] or not bank_cfgs[fname]['credit']: return messagebox.showwarning("错误", f"{fname} 配置不全")

        final_map = {}
        for gl, cb in self.mapping_combos:
            v = cb.get()
            if v != "(未匹配)":
                full = next((x['path'] for x in self.engine.bank_files_info if x['name'] == v), None)
                if full: final_map[gl] = full
        
        if not final_map: return messagebox.showwarning("提示", "无有效映射")
        
        stg = {'aggregation': self.var_hungarian.get(), 'subset': self.var_subset.get(), 'name_threshold': self.slider_thresh.get()}
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="审计核对底稿.xlsx")
        if not out: return

        # 申请中断信号
        stop_event = None
        if hasattr(self, 'app'): stop_event = self.app.register_task(self.module_index)

        self.btn_start_match.configure(state="disabled", text="核对中...")
        def t():
            try:
                # 传入 stop_event
                ok, msg = self.engine.execute_reconciliation(final_map, gl_cfg, bank_cfgs, out, stg, stop_event)
                self.log(msg); messagebox.showinfo("完成", "核对结束")
            except Exception as e:
                import traceback; traceback.print_exc(); self.log(f"Error: {e}")
            finally: 
                # 销假
                if hasattr(self, 'app'): self.app.finish_task(self.module_index)
                self.btn_start_match.configure(state="normal", text="🚀 生成审计底稿")
        threading.Thread(target=t, daemon=True).start()