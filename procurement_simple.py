"""
期货原料最佳采购优化 - 纯NumPy实现（无需scipy）
使用梯度下降求解带约束的优化问题
"""

import numpy as np
import matplotlib.pyplot as plt


class SimpleProcurementOptimizer:
    """简化版采购优化器"""
    
    def __init__(self, 
                 holding_cost=10.0,      # 库存持有成本（元/吨/月）
                 shortage_cost=100.0,    # 缺货成本（元/吨）
                 ordering_cost=500.0,    # 订货固定成本（元）
                 max_inventory=200.0,    # 最大库存
                 initial_inventory=50.0  # 初始库存
                 ):
        self.holding_cost = holding_cost
        self.shortage_cost = shortage_cost
        self.ordering_cost = ordering_cost
        self.max_inventory = max_inventory
        self.initial_inventory = initial_inventory
    
    def calculate_cost(self, orders, prices, demands):
        """
        计算给定采购策略的总成本
        
        orders: 各月订购量数组
        prices: 各月价格数组
        demands: 各月需求量数组
        """
        T = len(orders)
        inventory = self.initial_inventory
        total_cost = 0.0
        
        for t in range(T):
            # 采购成本
            if orders[t] > 0:
                total_cost += self.ordering_cost + prices[t] * orders[t]
            
            # 更新库存
            inventory += orders[t]
            
            # 满足需求
            if inventory >= demands[t]:
                inventory -= demands[t]
            else:
                # 缺货惩罚
                shortage = demands[t] - inventory
                total_cost += self.shortage_cost * shortage
                inventory = 0
            
            # 库存不能超过容量
            inventory = min(inventory, self.max_inventory)
            
            # 库存持有成本
            total_cost += self.holding_cost * inventory
        
        return total_cost
    
    def optimize_gradient(self, prices, demands, learning_rate=0.1, n_iterations=1000):
        """
        使用投影梯度下降优化采购计划
        """
        T = len(prices)
        orders = np.zeros(T)
        
        for _ in range(n_iterations):
            # 数值计算梯度
            grad = np.zeros(T)
            eps = 0.01
            
            for t in range(T):
                orders_plus = orders.copy()
                orders_plus[t] += eps
                cost_plus = self.calculate_cost(orders_plus, prices, demands)
                cost_orig = self.calculate_cost(orders, prices, demands)
                grad[t] = (cost_plus - cost_orig) / eps
            
            # 梯度下降更新
            orders -= learning_rate * grad
            
            # 投影到可行域（非负约束）
            orders = np.maximum(orders, 0)
            
            # 投影到库存约束
            inventory = self.initial_inventory
            for t in range(T):
                max_order = self.max_inventory - inventory
                orders[t] = min(orders[t], max(0, max_order))
                inventory += orders[t] - demands[t]
                inventory = max(0, min(inventory, self.max_inventory))
        
        return orders
    
    def optimize_heuristic(self, prices, demands, safety_stock=20):
        """
        启发式策略：(s, S) 策略的简化版
        当库存低于再订购点时订购至目标库存
        """
        T = len(prices)
        orders = np.zeros(T)
        inventory = self.initial_inventory
        
        for t in range(T):
            # 计算目标库存（基于未来需求预测）
            future_demand = sum(demands[t:min(t+3, T)])  # 看未来3个月
            target_inventory = future_demand + safety_stock
            
            # 价格因素：价格高时减少订购
            avg_price = np.mean(prices)
            if prices[t] > avg_price * 1.1:  # 高价
                target_inventory *= 0.8
            elif prices[t] < avg_price * 0.9:  # 低价
                target_inventory *= 1.2
            
            # 计算订购量
            order_needed = target_inventory - inventory
            
            # 检查约束
            max_order = self.max_inventory - inventory
            orders[t] = max(0, min(order_needed, max_order))
            
            # 更新库存状态
            inventory += orders[t]
            inventory = max(0, inventory - demands[t])
            inventory = min(inventory, self.max_inventory)
        
        return orders


def simulate_price_and_demand(n_periods=6, n_scenarios=100, seed=42):
    """
    模拟价格路径（几何布朗运动）和需求
    """
    np.random.seed(seed)
    
    current_price = 5000.0
    drift = 0.05      # 年化趋势5%
    volatility = 0.20 # 年化波动率20%
    dt = 1/12         # 月时间步长
    
    price_scenarios = np.zeros((n_scenarios, n_periods))
    price_scenarios[:, 0] = current_price
    
    for t in range(1, n_periods):
        dW = np.random.standard_normal(n_scenarios) * np.sqrt(dt)
        price_change = (drift - 0.5 * volatility**2) * dt + volatility * dW
        price_scenarios[:, t] = price_scenarios[:, t-1] * np.exp(price_change)
    
    # 需求：正态分布，均值100吨/月，标准差20吨
    demand_scenarios = np.maximum(0, np.random.normal(100, 20, (n_scenarios, n_periods)))
    
    return price_scenarios, demand_scenarios


def demo():
    """演示优化"""
    print("=" * 70)
    print("期货原料采购优化示例 (纯NumPy实现)")
    print("=" * 70)
    
    # 生成情景
    n_periods = 6
    n_scenarios = 100
    price_scenarios, demand_scenarios = simulate_price_and_demand(n_periods, n_scenarios)
    
    # 使用第一个情景作为基准
    prices = price_scenarios[0]
    demands = demand_scenarios[0]
    
    print(f"\n规划期: {n_periods}个月")
    print(f"初始库存: 50吨")
    print(f"最大库存: 200吨")
    print(f"当前价格: {prices[0]:.2f}元/吨")
    
    optimizer = SimpleProcurementOptimizer()
    
    # 方法1：启发式优化
    print("\n" + "-" * 50)
    print("【启发式策略优化】")
    print("-" * 50)
    
    orders_heuristic = optimizer.optimize_heuristic(prices, demands)
    cost_heuristic = optimizer.calculate_cost(orders_heuristic, prices, demands)
    
    print(f"总成本: {cost_heuristic:,.2f}元\n")
    print("月份 | 价格(元/吨) | 需求(吨) | 订购(吨) | 备注")
    print("-" * 60)
    
    inventory = 50.0
    for t in range(n_periods):
        note = ""
        avg_price = np.mean(prices)
        if prices[t] > avg_price * 1.1:
            note = "高价区间，减少订购"
        elif prices[t] < avg_price * 0.9:
            note = "低价区间，增加订购"
        
        print(f"{t+1:2d}   | {prices[t]:10.2f} | {demands[t]:8.2f} | {orders_heuristic[t]:8.2f} | {note}")
        
        inventory += orders_heuristic[t]
        inventory = max(0, inventory - demands[t])
    
    # 方法2：梯度下降优化
    print("\n" + "-" * 50)
    print("【梯度下降优化】")
    print("-" * 50)
    
    orders_grad = optimizer.optimize_gradient(prices, demands, learning_rate=0.5, n_iterations=500)
    cost_grad = optimizer.calculate_cost(orders_grad, prices, demands)
    
    print(f"总成本: {cost_grad:,.2f}元")
    print(f"成本节省: {((cost_heuristic - cost_grad) / cost_heuristic * 100):.2f}%\n")
    
    # 蒙特卡洛模拟评估策略
    print("-" * 50)
    print("【蒙特卡洛模拟评估】(100个情景)")
    print("-" * 50)
    
    costs = []
    for i in range(n_scenarios):
        orders = optimizer.optimize_heuristic(price_scenarios[i], demand_scenarios[i])
        cost = optimizer.calculate_cost(orders, price_scenarios[i], demand_scenarios[i])
        costs.append(cost)
    
    costs = np.array(costs)
    print(f"平均成本: {np.mean(costs):,.2f}元")
    print(f"成本标准差: {np.std(costs):,.2f}元")
    print(f"最小成本: {np.min(costs):,.2f}元")
    print(f"最大成本: {np.max(costs):,.2f}元")
    print(f"95%置信区间: [{np.percentile(costs, 2.5):,.2f}, {np.percentile(costs, 97.5):,.2f}]元")
    
    # 可视化
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    # 价格路径
    ax = axes[0, 0]
    for i in range(min(20, n_scenarios)):
        ax.plot(range(1, n_periods+1), price_scenarios[i], alpha=0.3, color='blue')
    ax.plot(range(1, n_periods+1), prices, 'r-', linewidth=2, label='基准情景')
    ax.set_xlabel('月份')
    ax.set_ylabel('价格（元/吨）')
    ax.set_title('期货价格情景模拟（几何布朗运动）')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    # 采购计划对比
    ax = axes[0, 1]
    x = np.arange(1, n_periods+1)
    width = 0.35
    ax.bar(x - width/2, orders_heuristic, width, label='启发式策略', alpha=0.8, color='steelblue')
    ax.bar(x + width/2, orders_grad, width, label='梯度优化', alpha=0.8, color='coral')
    ax.plot(x, demands, 'g-o', label='需求量', linewidth=2, markersize=8)
    ax.set_xlabel('月份')
    ax.set_ylabel('数量（吨）')
    ax.set_title('采购计划对比')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    # 库存变化
    ax = axes[1, 0]
    inventory_hist = [50.0]
    inv = 50.0
    for t in range(n_periods):
        inv += orders_heuristic[t]
        inv = max(0, inv - demands[t])
        inv = min(inv, 200.0)
        inventory_hist.append(inv)
    
    ax.plot(range(n_periods+1), inventory_hist, 'g-o', linewidth=2, markersize=8)
    ax.axhline(y=200, color='r', linestyle='--', label='最大库存')
    ax.fill_between(range(n_periods+1), 0, inventory_hist, alpha=0.3, color='green')
    ax.set_xlabel('月份')
    ax.set_ylabel('库存（吨）')
    ax.set_title('库存水平变化（启发式策略）')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    # 成本分布
    ax = axes[1, 1]
    ax.hist(costs, bins=20, alpha=0.7, color='orange', edgecolor='black')
    ax.axvline(np.mean(costs), color='r', linestyle='--', linewidth=2, 
               label=f'平均: {np.mean(costs):,.0f}')
    ax.axvline(np.median(costs), color='b', linestyle='--', linewidth=2,
               label=f'中位数: {np.median(costs):,.0f}')
    ax.set_xlabel('总成本（元）')
    ax.set_ylabel('频次')
    ax.set_title('成本分布（蒙特卡洛模拟）')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig('/root/.openclaw/workspace/procurement_optimization.png', dpi=150, bbox_inches='tight')
    print(f"\n图表已保存: procurement_optimization.png")
    
    return {
        'orders_heuristic': orders_heuristic,
        'orders_gradient': orders_grad,
        'costs_mc': costs
    }


if __name__ == '__main__':
    results = demo()
