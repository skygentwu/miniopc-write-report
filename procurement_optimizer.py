"""
期货原料最佳采购策略优化模型
目标：在价格波动和需求不确定下，确定最佳采购时间、数量和价格区间
方法：情景随机规划 + 滚动优化
"""

import numpy as np
import pandas as pd
from scipy.optimize import minimize, linprog
from scipy.stats import norm
import matplotlib.pyplot as plt
from dataclasses import dataclass
from typing import List, Tuple
import warnings
warnings.filterwarnings('ignore')


@dataclass
class ProcurementParams:
    """采购参数配置"""
    # 时间参数
    planning_horizon: int = 12  # 规划期（月）
    
    # 成本参数
    holding_cost: float = 5.0      # 单位库存持有成本（元/吨/月）
    shortage_cost: float = 50.0    # 单位缺货成本（元/吨）
    ordering_cost: float = 1000.0  # 每次订货固定成本（元）
    
    # 库存参数
    initial_inventory: float = 100.0  # 初始库存（吨）
    max_inventory: float = 500.0      # 最大库存容量（吨）
    
    # 资金约束
    budget: float = 1000000.0  # 总预算（元）
    
    # 期货参数
    futures_premium: float = 0.05  # 期货升水/贴水比例
    margin_rate: float = 0.10      # 保证金比例


class PriceSimulator:
    """价格情景模拟器 - 基于几何布朗运动"""
    
    def __init__(self, current_price: float, drift: float, volatility: float):
        self.current_price = current_price
        self.drift = drift          # 价格趋势（年化）
        self.volatility = volatility  # 价格波动率（年化）
    
    def generate_scenarios(self, n_periods: int, n_scenarios: int, dt: float = 1/12) -> np.ndarray:
        """
        生成价格情景矩阵 (n_scenarios x n_periods)
        使用几何布朗运动模拟价格路径
        """
        scenarios = np.zeros((n_scenarios, n_periods))
        scenarios[:, 0] = self.current_price
        
        for t in range(1, n_periods):
            # dS/S = mu*dt + sigma*dW
            dW = np.random.standard_normal(n_scenarios) * np.sqrt(dt)
            price_change = (self.drift - 0.5 * self.volatility**2) * dt + self.volatility * dW
            scenarios[:, t] = scenarios[:, t-1] * np.exp(price_change)
        
        return scenarios


class ProcurementOptimizer:
    """采购优化器"""
    
    def __init__(self, params: ProcurementParams):
        self.params = params
    
    def deterministic_optimize(self, 
                              prices: np.ndarray, 
                              demands: np.ndarray) -> dict:
        """
        确定性优化（已知未来价格和需求）
        使用线性规划求解
        
        决策变量：[q1, q2, ..., qT, I1, I2, ..., IT, S1, S2, ..., ST]
        其中 q=订购量, I=库存, S=缺货量
        """
        T = len(prices)
        
        # 变量顺序：订购量(T) + 库存(T) + 缺货(T)
        n_vars = 3 * T
        
        # 目标函数：最小化总成本
        # 成本 = 采购成本 + 订货成本 + 库存成本 + 缺货成本
        c = np.zeros(n_vars)
        for t in range(T):
            c[t] = prices[t]          # 采购成本
            c[T + t] = self.params.holding_cost  # 库存成本
            c[2*T + t] = self.params.shortage_cost  # 缺货成本
        
        # 约束条件：A_ub @ x <= b_ub, A_eq @ x = b_eq
        A_eq = []
        b_eq = []
        
        # 库存平衡约束：I_t = I_{t-1} + q_t - d_t + S_t
        for t in range(T):
            row = np.zeros(n_vars)
            row[t] = 1                    # q_t
            row[T + t] = -1               # -I_t
            if t > 0:
                row[T + t - 1] = 1        # I_{t-1}
            row[2*T + t] = 1              # S_t
            A_eq.append(row)
            b_eq.append(-demands[t] if t > 0 else demands[t] - self.params.initial_inventory)
        
        A_eq = np.array(A_eq)
        b_eq = np.array(b_eq)
        
        # 不等式约束
        A_ub = []
        b_ub = []
        
        # 库存容量约束
        for t in range(T):
            row = np.zeros(n_vars)
            row[T + t] = 1
            A_ub.append(row)
            b_ub.append(self.params.max_inventory)
        
        # 预算约束（简化版，只考虑采购成本）
        row = np.zeros(n_vars)
        row[:T] = prices
        A_ub.append(row)
        b_ub.append(self.params.budget)
        
        A_ub = np.array(A_ub)
        b_ub = np.array(b_ub)
        
        # 变量边界
        bounds = [(0, None)] * T + [(0, self.params.max_inventory)] * T + [(0, None)] * T
        
        # 求解
        result = linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, 
                        bounds=bounds, method='highs')
        
        if result.success:
            return {
                'success': True,
                'orders': result.x[:T],
                'inventory': result.x[T:2*T],
                'shortage': result.x[2*T:],
                'total_cost': result.fun,
                'message': result.message
            }
        else:
            return {'success': False, 'message': result.message}
    
    def stochastic_optimize(self, 
                           price_scenarios: np.ndarray, 
                           demand_scenarios: np.ndarray,
                           scenario_weights: np.ndarray = None) -> dict:
        """
        随机规划优化（考虑多个价格/需求情景）
        
        这里采用简化策略：对每个情景求解，然后取期望
        """
        n_scenarios = price_scenarios.shape[0]
        if scenario_weights is None:
            scenario_weights = np.ones(n_scenarios) / n_scenarios
        
        all_results = []
        expected_orders = np.zeros(price_scenarios.shape[1])
        expected_cost = 0.0
        
        for i in range(n_scenarios):
            result = self.deterministic_optimize(
                price_scenarios[i], 
                demand_scenarios[i]
            )
            if result['success']:
                all_results.append(result)
                expected_orders += scenario_weights[i] * result['orders']
                expected_cost += scenario_weights[i] * result['total_cost']
        
        return {
            'success': len(all_results) > 0,
            'expected_orders': expected_orders,
            'expected_cost': expected_cost,
            'scenario_results': all_results,
            'n_scenarios_solved': len(all_results)
        }
    
    def policy_optimization(self, 
                           price_thresholds: List[Tuple[float, float]], 
                           quantity_rules: List[float],
                           price_simulator: PriceSimulator,
                           demand_mean: float,
                           demand_std: float,
                           n_simulations: int = 1000) -> dict:
        """
        策略参数优化（基于策略搜索）
        
        策略规则：
        - 当价格落入区间 [low, high] 时，订购 quantity
        - 可以定义多个价格区间和对应订购量
        """
        total_costs = []
        
        for _ in range(n_simulations):
            # 模拟价格和需求
            prices = price_simulator.generate_scenarios(
                self.params.planning_horizon, 1
            )[0]
            demands = np.maximum(0, np.random.normal(
                demand_mean, demand_std, self.params.planning_horizon
            ))
            
            # 按策略执行
            inventory = self.params.initial_inventory
            total_cost = 0.0
            
            for t in range(self.params.planning_horizon):
                price = prices[t]
                demand = demands[t]
                
                # 根据价格区间确定订购量
                order_qty = 0.0
                for (low, high), qty in zip(price_thresholds, quantity_rules):
                    if low <= price <= high:
                        order_qty = qty
                        break
                
                # 检查库存约束
                if inventory + order_qty > self.params.max_inventory:
                    order_qty = max(0, self.params.max_inventory - inventory)
                
                # 执行订购
                if order_qty > 0:
                    total_cost += self.params.ordering_cost + price * order_qty
                    inventory += order_qty
                
                # 满足需求
                if inventory >= demand:
                    inventory -= demand
                else:
                    # 缺货
                    shortage = demand - inventory
                    total_cost += self.params.shortage_cost * shortage
                    inventory = 0
                
                # 库存持有成本
                total_cost += self.params.holding_cost * inventory
            
            total_costs.append(total_cost)
        
        return {
            'mean_cost': np.mean(total_costs),
            'std_cost': np.std(total_costs),
            'percentile_95': np.percentile(total_costs, 95),
            'percentile_5': np.percentile(total_costs, 5)
        }


def demo():
    """演示案例"""
    print("=" * 60)
    print("期货原料采购优化示例")
    print("=" * 60)
    
    # 参数设置
    params = ProcurementParams(
        planning_horizon=6,      # 6个月规划期
        holding_cost=10.0,       # 库存成本10元/吨/月
        shortage_cost=100.0,     # 缺货成本100元/吨
        ordering_cost=500.0,     # 订货固定成本500元
        initial_inventory=50.0,  # 初始库存50吨
        max_inventory=200.0,     # 最大库存200吨
        budget=500000.0          # 预算50万元
    )
    
    # 价格模拟器（当前价格5000元/吨，年化趋势5%，波动率20%）
    price_sim = PriceSimulator(
        current_price=5000.0,
        drift=0.05,
        volatility=0.20
    )
    
    # 生成价格情景
    np.random.seed(42)
    price_scenarios = price_sim.generate_scenarios(
        n_periods=params.planning_horizon,
        n_scenarios=100
    )
    
    # 需求情景（均值100吨/月，标准差20吨）
    demand_scenarios = np.maximum(0, np.random.normal(
        100, 20, (100, params.planning_horizon)
    ))
    
    print(f"\n规划期: {params.planning_horizon}个月")
    print(f"初始库存: {params.initial_inventory}吨")
    print(f"当前现货价格: {price_sim.current_price}元/吨")
    print(f"价格情景数: {price_scenarios.shape[0]}")
    
    # 优化器
    optimizer = ProcurementOptimizer(params)
    
    # 1. 单情景确定性优化（使用平均价格）
    print("\n" + "-" * 40)
    print("【确定性优化】基于平均价格情景")
    print("-" * 40)
    
    avg_prices = np.mean(price_scenarios, axis=0)
    avg_demands = np.mean(demand_scenarios, axis=0)
    
    det_result = optimizer.deterministic_optimize(avg_prices, avg_demands)
    
    if det_result['success']:
        print(f"总成本: {det_result['total_cost']:,.2f}元")
        print(f"\n各月采购计划:")
        for t, (q, p, d) in enumerate(zip(det_result['orders'], avg_prices, avg_demands)):
            print(f"  第{t+1}月: 订购{q:6.2f}吨 @ {p:7.2f}元/吨 | 需求{d:6.2f}吨 | 期末库存{det_result['inventory'][t]:6.2f}吨")
    
    # 2. 随机规划优化
    print("\n" + "-" * 40)
    print("【随机规划优化】基于100个价格/需求情景")
    print("-" * 40)
    
    stoch_result = optimizer.stochastic_optimize(
        price_scenarios, demand_scenarios
    )
    
    if stoch_result['success']:
        print(f"期望总成本: {stoch_result['expected_cost']:,.2f}元")
        print(f"\n期望采购计划（各月平均）:")
        for t, q in enumerate(stoch_result['expected_orders']):
            avg_p = np.mean(price_scenarios[:, t])
            print(f"  第{t+1}月: 期望订购{q:6.2f}吨 @ 期望价格{avg_p:7.2f}元/吨")
    
    # 3. 策略优化示例
    print("\n" + "-" * 40)
    print("【策略优化】基于价格触发规则")
    print("-" * 40)
    
    # 定义策略：价格区间 -> 订购量
    price_thresholds = [
        (0, 4800),      # 低价区间：多买
        (4800, 5200),   # 中价区间：按需
        (5200, float('inf'))  # 高价区间：少买
    ]
    quantity_rules = [150.0, 100.0, 50.0]  # 对应订购量
    
    policy_result = optimizer.policy_optimization(
        price_thresholds, quantity_rules, price_sim,
        demand_mean=100.0, demand_std=20.0,
        n_simulations=500
    )
    
    print(f"策略规则：")
    for (low, high), qty in zip(price_thresholds, quantity_rules):
        print(f"  价格[{low:6.0f}, {high:6.0f}]元 → 订购{qty:5.1f}吨")
    
    print(f"\n蒙特卡洛模拟结果（500次）:")
    print(f"  平均总成本: {policy_result['mean_cost']:,.2f}元")
    print(f"  成本标准差: {policy_result['std_cost']:,.2f}元")
    print(f"  95%分位成本: {policy_result['percentile_95']:,.2f}元")
    print(f"  5%分位成本: {policy_result['percentile_5']:,.2f}元")
    
    # 可视化
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    # 价格情景
    ax = axes[0, 0]
    for i in range(min(20, len(price_scenarios))):
        ax.plot(range(1, params.planning_horizon+1), price_scenarios[i], 
                alpha=0.3, color='blue')
    ax.plot(range(1, params.planning_horizon+1), avg_prices, 'r-', linewidth=2, label='平均价格')
    ax.set_xlabel('月份')
    ax.set_ylabel('价格（元/吨）')
    ax.set_title('价格情景模拟')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    # 确定性优化结果
    ax = axes[0, 1]
    x = range(1, params.planning_horizon+1)
    ax.bar(x, det_result['orders'], alpha=0.7, label='订购量', color='steelblue')
    ax.plot(x, avg_demands, 'r-o', label='预期需求', linewidth=2)
    ax.set_xlabel('月份')
    ax.set_ylabel('数量（吨）')
    ax.set_title('确定性优化：订购计划 vs 需求')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    # 库存变化
    ax = axes[1, 0]
    ax.plot(x, det_result['inventory'], 'g-o', linewidth=2, markersize=8)
    ax.axhline(y=params.max_inventory, color='r', linestyle='--', label='最大库存')
    ax.fill_between(x, 0, det_result['inventory'], alpha=0.3, color='green')
    ax.set_xlabel('月份')
    ax.set_ylabel('库存（吨）')
    ax.set_title('库存水平变化')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    # 成本分布
    ax = axes[1, 1]
    # 重新运行模拟获取成本分布
    costs = []
    for _ in range(500):
        prices = price_sim.generate_scenarios(params.planning_horizon, 1)[0]
        demands = np.maximum(0, np.random.normal(100, 20, params.planning_horizon))
        
        inventory = params.initial_inventory
        total_cost = 0.0
        for t in range(params.planning_horizon):
            # 简单策略：按需订购
            order = max(0, demands[t] - inventory + 20)  # 安全库存
            if order > 0:
                total_cost += params.ordering_cost + prices[t] * order
                inventory += order
            inventory = max(0, inventory - demands[t])
            total_cost += params.holding_cost * inventory
        costs.append(total_cost)
    
    ax.hist(costs, bins=30, alpha=0.7, color='orange', edgecolor='black')
    ax.axvline(np.mean(costs), color='r', linestyle='--', linewidth=2, label=f'平均成本: {np.mean(costs):,.0f}')
    ax.set_xlabel('总成本（元）')
    ax.set_ylabel('频次')
    ax.set_title('成本分布（蒙特卡洛模拟）')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig('procurement_optimization.png', dpi=150, bbox_inches='tight')
    print(f"\n图表已保存: procurement_optimization.png")
    
    return {
        'deterministic': det_result,
        'stochastic': stoch_result,
        'policy': policy_result
    }


if __name__ == '__main__':
    results = demo()
