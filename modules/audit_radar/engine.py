import torch
import torch.nn as nn
import torch.optim as optim
import numpy as np
from .model import AuditAutoEncoder

class AuditEngine:
    def __init__(self, processor, device=None):
        self.processor = processor
        self.device = device if device else torch.device('cuda' if torch.cuda.is_available() else 'cpu')
        self.model = None
        
    def train_model(self, cats, conts, epochs=100, lr=0.001, log_callback=None, stop_event=None):
        self.model = AuditAutoEncoder(
            num_cont=len(self.processor.cont_cols),
            cat_dims=self.processor.cat_dims,
            emb_dims=self.processor.emb_dims
        ).to(self.device)
        
        optimizer = optim.Adam(self.model.parameters(), lr=lr)
        criterion = nn.MSELoss()
        
        self.model.train()
        
        final_loss = 0.0
        
        for epoch in range(epochs):
            if stop_event and stop_event.is_set():
                if log_callback: log_callback("训练中断")
                return False, None
                
            optimizer.zero_grad()
            decoded, original = self.model(cats, conts)
            loss = criterion(decoded, original)
            loss.backward()
            optimizer.step()
            
            final_loss = loss.item()
            
            # === 【修改点】逻辑优化：每10轮打印一次，且最后一轮必须打印 ===
            current_round = epoch + 1
            if log_callback and (current_round % 10 == 0 or current_round == epochs):
                log_callback(f"Training Epoch {current_round}/{epochs} | Loss: {final_loss:.4f}")
        
        return True, final_loss

    def predict_with_reason(self, cats, conts, raw_df=None, amt_cols=None, threshold=0):
        self.model.eval()
        with torch.no_grad():
            decoded, original = self.model(cats, conts)
            diff_square = (decoded - original) ** 2
            
            # 归因分析
            num_cont = len(self.processor.cont_cols)
            split_idx = original.shape[1] - num_cont
            
            if split_idx > 0:
                diff_cat = diff_square[:, :split_idx].mean(dim=1)
            else:
                diff_cat = torch.zeros(diff_square.shape[0]).to(self.device)
                
            if num_cont > 0:
                diff_cont = diff_square[:, split_idx:].mean(dim=1)
            else:
                diff_cont = torch.zeros(diff_square.shape[0]).to(self.device)
            
            scores = diff_square.mean(dim=1)
            
            reasons = []
            v_cat = diff_cat.cpu().numpy()
            v_cont = diff_cont.cpu().numpy()
            
            for c, a in zip(v_cat, v_cont):
                if a > c: reasons.append("金额异常")
                else: reasons.append("科目/组合模式异常")

            scores = scores.cpu().numpy()

        # 重要性水平过滤逻辑 (此处仅做计算，统计逻辑放到UI层展示更灵活)
        if threshold > 0 and raw_df is not None and amt_cols:
            max_abs_amt = raw_df[amt_cols].abs().max(axis=1).fillna(0).values
            mask_small = max_abs_amt < threshold
            scores[mask_small] = 0.0
            for i in range(len(reasons)):
                if mask_small[i]: reasons[i] = "忽略(金额小)"
            
        return scores, reasons