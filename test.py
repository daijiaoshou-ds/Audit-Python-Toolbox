import pandas as pd
import numpy as np
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler
import matplotlib.pyplot as plt

# 1. 设置中文显示
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

def create_mock_data():
    """生成模拟财务数据：47家正常，3家异常"""
    np.random.seed(42)
    companies = [f"分公司_{i}" for i in range(1, 51)]
    
    # 生成正常数据 (收入和成本正相关)
    revenue = np.random.normal(1000, 200, 50)
    cost = revenue * 0.6 + np.random.normal(0, 50, 50) # 正常毛利 40%
    expense = revenue * 0.2 + np.random.normal(0, 20, 50) # 正常费用 20%
    
    data = pd.DataFrame({
        '公司名称': companies,
        '主营收入': revenue,
        '主营成本': cost,
        '管理费用': expense
    })

    # === 造假：制造 3 个异常样本 ===
    
    # 异常1：收入很高，成本极低 (虚增利润)
    data.loc[0, '主营收入'] = 2000
    data.loc[0, '主营成本'] = 400  # 毛利高达 80%，极度异常
    data.loc[0, '公司名称'] = "异常公司_A(高毛利)"

    # 异常2：收入正常，费用极高 (利益输送?)
    data.loc[1, '管理费用'] = 800  # 费用率爆表
    data.loc[1, '公司名称'] = "异常公司_B(高费用)"
    
    # 异常3：数据看起来都在范围内，但组合很怪
    # 比如收入很低，但费用很高
    data.loc[2, '主营收入'] = 500
    data.loc[2, '管理费用'] = 400
    data.loc[2, '公司名称'] = "异常公司_C(低收高费)"

    return data

def run_analysis():
    print("正在生成模拟数据...")
    df = create_mock_data()
    
    # 2. 特征工程：计算比率 (这一步是核心！)
    # 我们只分析比率，不分析绝对金额
    print("正在计算财务比率...")
    df_ratios = pd.DataFrame()
    df_ratios['公司名称'] = df['公司名称']
    
    # 避免除以0
    df_ratios['毛利率'] = (df['主营收入'] - df['主营成本']) / df['主营收入']
    df_ratios['费用率'] = df['管理费用'] / df['主营收入']
    df_ratios['成本收入比'] = df['主营成本'] / df['主营收入']
    
    # 准备矩阵
    features = ['毛利率', '费用率', '成本收入比']
    x = df_ratios[features].values
    
    # 3. 数据标准化 (Mean=0, Std=1)
    x = StandardScaler().fit_transform(x)
    
    # 4. PCA 降维 (降到2维以便画图)
    pca = PCA(n_components=2)
    principalComponents = pca.fit_transform(x)
    
    # 生成结果表
    result_df = pd.DataFrame(data=principalComponents, columns=['PC1', 'PC2'])
    result_df['公司名称'] = df['公司名称']
    
    # 计算每个点距离圆心(0,0)的距离 -> 这就是“异常程度”
    result_df['异常得分'] = np.sqrt(result_df['PC1']**2 + result_df['PC2']**2)
    
    # 按异常程度排序
    result_df = result_df.sort_values(by='异常得分', ascending=False)
    
    print("\n=== 异常检测结果 (前5名) ===")
    print(result_df[['公司名称', '异常得分']].head(5))

    # 5. 画图
    plt.figure(figsize=(10, 8))
    
    # 画所有点
    plt.scatter(result_df['PC1'], result_df['PC2'], c='blue', alpha=0.5, label='正常样本')
    
    # 标记前 5 个异常点为红色
    top_5 = result_df.head(5)
    plt.scatter(top_5['PC1'], top_5['PC2'], c='red', s=100, label='高风险样本')
    
    # 给前 5 名标上名字
    for i, row in top_5.iterrows():
        plt.text(row['PC1']+0.1, row['PC2']+0.1, row['公司名称'], fontsize=9)
        
    plt.title('智能审计雷达：基于 PCA 的多维异常检测', fontsize=16)
    plt.xlabel(f'主成分 1 (解释度 {pca.explained_variance_ratio_[0]:.2%})')
    plt.ylabel(f'主成分 2 (解释度 {pca.explained_variance_ratio_[1]:.2%})')
    plt.axhline(y=0, color='k', linestyle='--', alpha=0.3)
    plt.axvline(x=0, color='k', linestyle='--', alpha=0.3)
    plt.grid(True, alpha=0.3)
    plt.legend()
    
    print("\n正在显示图表... (请查看弹出的窗口)")
    plt.show()

if __name__ == "__main__":
    run_analysis()