import pandas as pd
import numpy as np
from sklearn.preprocessing import LabelEncoder, StandardScaler
import torch

class AuditDataProcessor:
    def __init__(self):
        self.label_encoders = {}
        self.scaler = StandardScaler()
        self.cat_cols = []
        self.cont_cols = []
        self.cat_dims = []
        self.emb_dims = []
        
    def preprocess(self, df, cont_col_names, cat_col_names):
        """
        df: 原始 DataFrame
        cont_col_names: 金额列名列表
        cat_col_names: 分类列名列表
        """
        # 拷贝副本
        data = df.copy()
        
        # --- 1. 智能日期特征提取 (新增功能) ---
        # 我们会在 cat_col_names 里寻找像 '日期'、'Date'、'Time' 这样的列
        # 如果找到了，就自动生成 'Month' 和 'Is_Weekend' 特征
        final_cat_cols = []
        
        for col in cat_col_names:
            # 尝试转为日期格式
            is_date = False
            if '日期' in str(col) or 'time' in str(col).lower() or 'date' in str(col).lower():
                try:
                    # 尝试转换，失败则忽略
                    temp_series = pd.to_datetime(data[col], errors='coerce')
                    if temp_series.notna().sum() > len(data) * 0.5: # 超过一半转换成功
                        is_date = True
                        # 生成新特征
                        data[f'{col}_Month'] = temp_series.dt.month.fillna(0).astype(int).astype(str)
                        # data[f'{col}_Weekday'] = temp_series.dt.weekday.fillna(0).astype(int).astype(str) # 0-6
                        # 甚至可以加一个是否周末
                        data[f'{col}_IsWeekend'] = temp_series.dt.weekday.apply(lambda x: 'Yes' if x>=5 else 'No').astype(str)
                        
                        final_cat_cols.append(f'{col}_Month')
                        final_cat_cols.append(f'{col}_IsWeekend')
                except:
                    pass
            
            # 如果不是日期，或者转换失败，就作为普通分类处理
            # 但我们要过滤掉唯一值太多的列（比如摘要、凭证号），因为它们是噪音
            if not is_date:
                unique_count = data[col].nunique()
                # 如果唯一值数量超过行数的 80% (且行数>100)，说明这列几乎每行都不一样(如摘要、编号)
                # 这种列对于自编码器是毁灭性的噪音，必须丢弃！
                if len(data) > 100 and unique_count > len(data) * 0.8:
                    print(f"[警告] 列 '{col}' 的唯一值过多 ({unique_count})，判定为噪音，已自动跳过。")
                    continue
                
                final_cat_cols.append(col)

        self.cat_cols = final_cat_cols
        self.cont_cols = cont_col_names
        
        # --- 2. 处理连续变量 (金额) ---
        processed_conts = []
        for col in cont_col_names:
            # 转数字，非数字变0
            data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0)
            
            # 关键：保留正负号信息的 Log 处理
            # log1p(abs(x)) * sign(x)
            # 这样 AI 既能理解金额大小，也能理解借贷方向
            col_val = data[col].values
            sign = np.sign(col_val)
            log_abs = np.log1p(np.abs(col_val))
            transformed = log_abs * sign
            
            processed_conts.append(transformed.reshape(-1, 1))
            
        if processed_conts:
            cont_matrix = np.hstack(processed_conts)
            self.cont_data = self.scaler.fit_transform(cont_matrix)
        else:
            self.cont_data = np.array([])

        # --- 3. 处理分类变量 ---
        self.cat_dims = []
        self.emb_dims = []
        
        for col in self.cat_cols:
            data[col] = data[col].fillna("Unknown").astype(str)
            le = LabelEncoder()
            data[col] = le.fit_transform(data[col])
            self.label_encoders[col] = le
            
            dim = len(le.classes_)
            self.cat_dims.append(dim)
            self.emb_dims.append(min(50, (dim + 1) // 2))
            
        return data

    def get_tensors(self, df, device):
        if self.cat_cols:
            cats = np.stack([df[c].values for c in self.cat_cols], 1)
            cats = torch.tensor(cats, dtype=torch.long).to(device)
        else:
            cats = torch.tensor([], dtype=torch.long).to(device)
        
        if self.cont_cols:
            conts = torch.tensor(self.cont_data, dtype=torch.float).to(device)
        else:
            conts = torch.tensor([], dtype=torch.float).to(device)
        
        return cats, conts