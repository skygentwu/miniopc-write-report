"""
铁矿石(r,Q)库存策略模型复现（纯NumPy实现，无需scipy）
基于：许贵斌, 赵旭, 李雪婷. 基于随机理论的我国铁矿石库存优化. 上海海事大学学报, 2013
"""

import numpy as np
import matplotlib.pyplot as plt


def norm_pdf(x, mu=0, sigma=1):
    """正态分布概率密度函数"""
    return np.exp(-0.5 * ((x - mu) / sigma) ** 2) / (sigma * np.sqrt(2 * np.pi))


def norm_cdf(x, mu=0, sigma=1):
    """正态分布累积分布函数（近似）"""
    # 使用误差函数近似
    z = (x - mu) / (sigma * np.sqrt(2))
    # 多项式近似（Abramowitz and Stegun）
    t = 1 / (1 + 0.3275911 * np.abs(z))
    y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * np.exp(-z * z)
    return 0.5 * (1 + np.sign(z) * y)


def norm_ppf(p, mu=0, sigma=1):
    """正态分布分位数函数（逆CDF，近似）"""
    # 使用有理近似（Peter J. Acklam）
    if p < 0 or p > 1:
        return np.nan
    
    a1 = -3.969683028665376e+01
    a2 = 2.209460984245205e+02
    a3 = -2.759285104469687e+02
    a4 = 1.383577518672690e+02
    a5 = -3.066479806614716e+01
    a6 = 2.506628277459239e+00
    
    b1 = -5.447609879822406e+01
    b2 = 1.615858368580409e+02
    b3 = -1.556989798598866e+02
    b4 = 6.680131188771972e+01
    b5 = -1.328068155288572e+01
    
    c1 = -7.784894002430293e-03
    c2 = -3.223964580411365e-01
    c3 = -2.400758277161838e+00
    c4 = -2.549732539343734e+00
    c5 = 4.374664141464968e+00
    c6 = 2.938163982698783e+00
    
    d1 = 7.784695709041462e-03
    d2 = 3.224671290700398e-01
    d3 = 2.445134137142996e+00
    d4 = 3.754408661907416e+00
    
    p_low = 0.02425
    p_high = 1 - p_low
    
    if p < p_low:
        q = np.sqrt(-2 * np.log(p))
        z = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / \
            ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    elif p <= p_high:
        q = p - 0.5
        r = q * q
        z = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / \
            (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
    else:
        q = np.sqrt(-2 * np.log(1 - p))
        z = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / \
             ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    
    return mu + sigma * z


class RQInventoryModel:
    """
    (r,Q)库存优化模型 - 纯NumPy实现
    """
    
    def __init__(self,
                 demand_mean=100,      # 月均需求（万吨）
                 demand_std=20,        # 需求标准差
                 lead_time=2,          # 提前期（月）
                 ordering_cost=500,    # 订货固定成本（万元/次）
                 holding_cost=10,      # 单位持有成本（元/吨/月）
                 shortage_cost=50,     # 单位缺货成本（元/吨）
                 purchase_price=750):  # 采购单价（元/吨）
        
        self.demand_mean = demand_mean
        self.demand_std = demand_std
        self.lead_time = lead_time
        self.ordering_cost = ordering_cost
        self.holding_cost = holding_cost
        self.shortage_cost = shortage_cost
        self.purchase_price = purchase_price
        
        # 提前期需求分布参数
        self.LT_mean = demand_mean * lead_time
        self.LT_std = demand_std * np.sqrt(lead_time)
        
    def calculate_expected_shortage(self, r):
        """计算期望缺货量 E[(D_LT - r)^+]"""
        z = (r - self.LT_mean) / self.LT_std
        # 标准正态的期望超出量 = φ(z) - z·(1-Φ(z))
        pdf_z = norm_pdf(z)
        cdf_z = norm_cdf(z)
        expected_excess = self.LT_std * (pdf_z - z * (1 - cdf_z))
        return max(0, expected_excess)
    
    def calculate_cycle_service_level(self, r):
        """计算周期服务水平 P(D_LT <= r)"""
        return norm_cdf(r, self.LT_mean, self.LT_std)
    
    def total_cost(self, Q, r):
        """计算单位时间（年）总成本"""
        annual_demand = self.demand_mean * 12
        
        # 订货成本
        ordering_cost_annual = (annual_demand / Q) * self.ordering_cost
        
        # 采购成本
        purchase_cost = annual_demand * self.purchase_price / 10000
        
        # 持有成本
        avg_inventory = Q / 2 + max(0, r - self.LT_mean)
        holding_cost_annual = avg_inventory * 12 * self.holding_cost / 10000
        
        # 缺货成本
        expected_shortage = self.calculate_expected_shortage(r)
        shortage_cost_annual = (annual_demand / Q) * expected_shortage * self.shortage_cost / 10000
        
        total = ordering_cost_annual + purchase_cost + holding_cost_annual + shortage_cost_annual
        
        return {
            'total': total,
            'ordering': ordering_cost_annual,
            'purchase': purchase_cost,
            'holding': holding_cost_annual,
            'shortage': shortage_cost_annual,
            'avg_inventory': avg_inventory,
            'expected_shortage': expected_shortage
        }
    
    def optimal_Q_analytical(self, r):
        """解析法求解最优Q（给定r）- 修正EOQ公式"""
        annual_demand = self.demand_mean * 12
        expected_shortage = self.calculate_expected_shortage(r)
        h_annual = self.holding_cost * 12 / 10000
        
        # 考虑缺货成本的修正EOQ
        effective_ordering_cost = self.ordering_cost + self.shortage_cost / 10000 * expected_shortage
        Q_opt = np.sqrt(2 * annual_demand * effective_ordering_cost / h_annual)
        return Q_opt
    
    def optimal_r_analytical(self, Q):
        """解析法求解最优r（给定Q）- 临界比率法"""
        annual_demand = self.demand_mean * 12
        
        # 缺货成本与持有成本的权衡
        shortage_cost_per_cycle = self.shortage_cost / 10000 * annual_demand / Q
        holding_cost_per_cycle = self.holding_cost * 12 / 10000 * Q
        
        # 临界比率
        critical_ratio = shortage_cost_per_cycle / (shortage_cost_per_cycle + holding_cost_per_cycle)
        
        # 正态分布的分位数
        r_opt = norm_ppf(critical_ratio, self.LT_mean, self.LT_std)
        return r_opt
    
    def solve_iterative(self, max_iter=100, tol=0.01, min_service_level=0.95):
        """迭代法求解最优(r, Q)
        
        参数:
            min_service_level: 最低服务水平约束，确保r不会过低
        """
        # 初始值：经典EOQ
        annual_demand = self.demand_mean * 12
        h_annual = self.holding_cost * 12 / 10000
        Q = np.sqrt(2 * annual_demand * self.ordering_cost / h_annual)
        
        # 初始r：确保至少满足最低服务水平
        r_min = norm_ppf(min_service_level, self.LT_mean, self.LT_std)
        r = max(self.LT_mean, r_min)
        
        print(f"迭代求解过程:")
        print(f"{'Iter':>4} | {'Q':>8} | {'r':>8} | {'Total Cost':>12} | {'Service Level':>13}")
        print("-" * 60)
        
        for i in range(max_iter):
            # 步骤1: 固定Q，优化r（但有最低服务水平约束）
            r_calc = self.optimal_r_analytical(Q)
            r_new = max(r_calc, r_min)  # 确保不低于最低服务水平
            
            # 步骤2: 固定r，优化Q
            Q_new = self.optimal_Q_analytical(r_new)
            
            cost = self.total_cost(Q_new, r_new)['total']
            service_level = self.calculate_cycle_service_level(r_new)
            
            if i < 5 or i % 10 == 0:
                print(f"{i+1:>4} | {Q_new:>8.2f} | {r_new:>8.2f} | {cost:>12.2f} | {service_level:>12.2%}")
            
            # 收敛判断
            if abs(Q - Q_new) < tol and abs(r - r_new) < tol:
                print(f"{i+1:>4} | {Q_new:>8.2f} | {r_new:>8.2f} | {cost:>12.2f} | {service_level:>12.2%}  **Converged**")
                break
            
            Q, r = Q_new, r_new
        
        return Q, r


def demo():
    """演示：钢铁企业铁矿石采购(r,Q)优化"""
    
    print("="*70)
    print("(r,Q) Inventory Model - Xu et al. (2013) Reproduction")
    print("="*70)
    
    # 模型参数（调整为更合理的规模）
    model = RQInventoryModel(
        demand_mean=10,       # 月均需求10万吨（更合理规模）
        demand_std=2,         # 标准差2万吨
        lead_time=2,          # 提前期2个月（巴西矿船期）
        ordering_cost=50,     # 订货固定成本50万元/次
        holding_cost=10,      # 库存成本10元/吨/月
        shortage_cost=200,    # 缺货成本200元/吨（高惩罚促使高服务水平）
        purchase_price=750    # 采购价750元/吨
    )
    
    print(f"\n[Model Parameters]")
    print(f"Monthly Demand: {model.demand_mean} (10k tons)")
    print(f"Demand Std: {model.demand_std} (10k tons)")
    print(f"Lead Time: {model.lead_time} months")
    print(f"LT Demand Mean: {model.LT_mean:.2f} (10k tons)")
    print(f"LT Demand Std: {model.LT_std:.2f} (10k tons)")
    
    # 求解最优(r,Q) - 使用95%最低服务水平
    Q_opt, r_opt = model.solve_iterative(min_service_level=0.95)
    
    # 详细成本分析
    print(f"\n[Optimal Solution]")
    print(f"Optimal Order Quantity Q*: {Q_opt:.2f} (10k tons)")
    print(f"Optimal Reorder Point r*: {r_opt:.2f} (10k tons)")
    
    cost_detail = model.total_cost(Q_opt, r_opt)
    print(f"\nAnnual Cost Breakdown:")
    print(f"  Ordering Cost:  {cost_detail['ordering']:>12.2f} (10k CNY)")
    print(f"  Purchase Cost:  {cost_detail['purchase']:>12.2f} (10k CNY)")
    print(f"  Holding Cost:   {cost_detail['holding']:>12.2f} (10k CNY)")
    print(f"  Shortage Cost:  {cost_detail['shortage']:>12.2f} (10k CNY)")
    print(f"  Total Cost:     {cost_detail['total']:>12.2f} (10k CNY)")
    
    print(f"\nOperational Metrics:")
    print(f"  Avg Inventory: {cost_detail['avg_inventory']:.2f} (10k tons)")
    print(f"  Expected Shortage: {cost_detail['expected_shortage']:.2f} (10k tons/cycle)")
    print(f"  Cycle Service Level: {model.calculate_cycle_service_level(r_opt):.2%}")
    
    # 可视化
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    # 1. (Q,r)成本热力图
    ax = axes[0, 0]
    Q_range = np.linspace(50, 300, 50)
    r_range = np.linspace(150, 300, 50)
    Q_grid, r_grid = np.meshgrid(Q_range, r_range)
    cost_grid = np.zeros_like(Q_grid)
    
    for i in range(len(r_range)):
        for j in range(len(Q_range)):
            cost_grid[i, j] = model.total_cost(Q_grid[i, j], r_grid[i, j])['total']
    
    contour = ax.contourf(Q_grid, r_grid, cost_grid, levels=20, cmap='viridis')
    ax.plot(Q_opt, r_opt, 'r*', markersize=15, label=f'Optimal (Q={Q_opt:.1f}, r={r_opt:.1f})')
    ax.set_xlabel('Q (Order Quantity)')
    ax.set_ylabel('r (Reorder Point)')
    ax.set_title('Total Cost Heatmap (10k CNY)')
    plt.colorbar(contour, ax=ax)
    ax.legend()
    
    # 2. 提前期需求分布
    ax = axes[0, 1]
    x = np.linspace(model.LT_mean - 3*model.LT_std, 
                    model.LT_mean + 3*model.LT_std, 200)
    y = [norm_pdf(xi, model.LT_mean, model.LT_std) for xi in x]
    ax.fill_between(x, y, alpha=0.3, label='Demand Distribution')
    ax.axvline(r_opt, color='r', linestyle='--', linewidth=2, label=f'Reorder Point r={r_opt:.1f}')
    ax.axvline(model.LT_mean, color='g', linestyle='--', label=f'Mean={model.LT_mean:.1f}')
    ax.fill_between(x[x > r_opt], [norm_pdf(xi, model.LT_mean, model.LT_std) for xi in x[x > r_opt]], 
                    alpha=0.3, color='red', label='Shortage Area')
    ax.set_xlabel('Demand During Lead Time (10k tons)')
    ax.set_ylabel('Probability Density')
    ax.set_title('Lead Time Demand Distribution')
    ax.legend()
    
    # 3. 成本构成对比
    ax = axes[1, 0]
    categories = ['Ordering', 'Purchase', 'Holding', 'Shortage']
    values = [cost_detail['ordering'], cost_detail['purchase'], 
              cost_detail['holding'], cost_detail['shortage']]
    colors = ['#3498db', '#95a5a6', '#2ecc71', '#e74c3c']
    bars = ax.bar(categories, values, color=colors, alpha=0.8, edgecolor='black')
    ax.set_ylabel('Cost (10k CNY)')
    ax.set_title('Annual Cost Breakdown')
    for i, v in enumerate(values):
        ax.text(i, v + max(values)*0.02, f'{v:.0f}', ha='center', fontsize=10)
    ax.grid(True, alpha=0.3, axis='y')
    
    # 4. Q和r的收敛过程
    ax = axes[1, 1]
    Q_history, r_history = [], []
    Q, r = np.sqrt(2 * 1200 * 500 / 0.12), model.LT_mean  # 初始值
    
    for _ in range(20):
        Q_history.append(Q)
        r_history.append(r)
        r = model.optimal_r_analytical(Q)
        Q = model.optimal_Q_analytical(r)
    
    ax.plot(range(1, len(Q_history)+1), Q_history, 'b-o', label='Q (Order Qty)', linewidth=2)
    ax.plot(range(1, len(r_history)+1), r_history, 'r-s', label='r (Reorder Point)', linewidth=2)
    ax.axhline(Q_opt, color='b', linestyle='--', alpha=0.5)
    ax.axhline(r_opt, color='r', linestyle='--', alpha=0.5)
    ax.set_xlabel('Iteration')
    ax.set_ylabel('Value (10k tons)')
    ax.set_title('Convergence of Q and r')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig('/root/.openclaw/workspace/rq_inventory_model.png', dpi=150)
    print(f"\nFigure saved: rq_inventory_model.png")
    
    return model, Q_opt, r_opt


if __name__ == '__main__':
    model, Q, r = demo()
