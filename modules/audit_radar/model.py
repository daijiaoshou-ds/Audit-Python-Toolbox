import torch
import torch.nn as nn

class AuditAutoEncoder(nn.Module):
    def __init__(self, num_cont, cat_dims, emb_dims):
        """
        num_cont: 连续特征数量 (如金额)
        cat_dims: 类别特征的去重数量列表 [科目数, 人员数...]
        emb_dims: 对应的嵌入向量维度 [10, 5...]
        """
        super().__init__()
        
        # 1. 嵌入层 (处理文字类特征)
        self.embeddings = nn.ModuleList([
            nn.Embedding(num, dim) for num, dim in zip(cat_dims, emb_dims)
        ])
        self.no_of_embs = sum(emb_dims)
        self.n_cont = num_cont
        
        # 总输入维度
        input_dim = self.no_of_embs + self.n_cont
        
        # 2. 编码器 (压缩)
        self.encoder = nn.Sequential(
            nn.Linear(input_dim, 64),
            nn.BatchNorm1d(64),
            nn.ReLU(),
            nn.Linear(64, 32),
            nn.BatchNorm1d(32),
            nn.ReLU(),
            nn.Linear(32, 8)  # 核心瓶颈层：只保留8个特征
        )
        
        # 3. 解码器 (还原)
        self.decoder = nn.Sequential(
            nn.Linear(8, 32),
            nn.BatchNorm1d(32),
            nn.ReLU(),
            nn.Linear(32, 64),
            nn.BatchNorm1d(64),
            nn.ReLU(),
            nn.Linear(64, input_dim)
        )

    def forward(self, x_cat, x_cont):
        # 处理嵌入
        x = [emb(x_cat[:, i]) for i, emb in enumerate(self.embeddings)]
        x = torch.cat(x, 1)
        
        # 拼接连续数据
        if self.n_cont > 0:
            x = torch.cat([x, x_cont], 1)
            
        encoded = self.encoder(x)
        decoded = self.decoder(encoded)
        return decoded, x  # 返回 (预测值, 真实值)