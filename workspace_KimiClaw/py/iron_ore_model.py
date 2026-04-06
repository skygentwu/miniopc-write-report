"""
铁矿石期货采购优化模型
目标：年度采购成本最小化
策略对比：保产量 vs 低库存
"""

import numpy as np
import matplotlib.pyplot as plt


class IronOreProcurementModel:
    """
    铁矿石采购优化模型
    
    铁矿石特性：
    - 期货合约：I2501, I2505, I2509（1月、5月、9月主力合约）
    - 品位：62%Fe 普氏指数定价
    - 到港周期：巴西约45天，澳洲约30天
    - 库存：港口库存+厂内库存
    """
    
    def __init__(self, 
                 annual_demand=12000,      # 年度需求（万吨）
                 initial_stock=800,        # 初始库存（万吨）
                 min_safety_stock=300,     # 安全库存（万吨）
                 max_stock=2000,           # 最大库存（万吨）
                 holding_cost=15,          # 库存成本（元/吨/月）
                 shortage_penalty=200,     # 缺货惩罚（元/吨）
                 transport_time=2):        # 运输时间（月）
        
        self.annual_demand = annual_demand
        self.initial_stock = initial_stock
        self.min_safety_stock = min_safety_stock
        self.max_stock = max_stock
        self.holding_cost = holding_cost
        self.shortage_penalty = shortage_penalty
        self.transport_time = transport_time
        
        # 月度需求（考虑季节性，钢厂春季复产、冬季限产）
        self.monthly_demand = self._generate_seasonal_demand()
        
    def _generate_seasonal_demand(self):
        """生成季节性需求（1-12月）"""
        base = self.annual_demand / 12
        # 季节性系数：3-5月旺季，7-8月淡季，11-12月冬储
        seasonal_factors = [0.9, 0.85, 1.1, 1.15, 1.1, 1.0, 
                           0.85, 0.9, 1.0, 1.05, 1.15, 1.15]
        return np.array([base * f for f in seasonal_factors])
    
    def simulate_prices(self, method='gbm', n_scenarios=100):
        """
        模拟铁矿石价格路径
        基准：普氏62%Fe指数，当前约100美元/干吨
        """
        np.random.seed(42)
        months = 12
        current_price = 750  # 人民币/吨（到岸价）
        
        if method == 'gbm':
            # 几何布朗运动
            drift = 0.02      # 年化2%趋势
            vol = 0.25        # 年化25%波动率
            dt = 1/12
            
            scenarios = np.zeros((n_scenarios, months))
            scenarios[:, 0] = current_price
            
            for t in range(1, months):
                dW = np.random.standard_normal(n_scenarios) * np.sqrt(dt)
                scenarios[:, t] = scenarios[:, t-1] * np.exp(
                    (drift - 0.5 * vol**2) * dt + vol * dW
                )
        
        elif method == 'historical':
            # 基于历史波动模拟
            scenarios = []
            for _ in range(n_scenarios):
                shocks = np.random.normal(0, 50, months)  # 月均波动50元
                price_path = current_price + np.cumsum(shocks)
                price_path = np.maximum(price_path, 400)  # 最低400元
                scenarios.append(price_path)
            scenarios = np.array(scenarios)
        
        return scenarios
    
    def strategy_max_production(self, prices, risk_aversion=0.5):
        """
        策略一：保产量最大化
        核心逻辑：宁可高库存，绝不缺料停产
        
        决策规则：
        - 库存 < 1.5倍安全库存 → 立即补库到最大
        - 不计较短期价格波动
        """
        months = 12
        orders = np.zeros(months)
        inventory = self.initial_stock
        total_cost = 0.0
        shortage_count = 0
        
        for t in range(months):
            # 保产量策略：库存低于1.5倍安全线就补
            if inventory < 1.5 * self.min_safety_stock:
                target = self.max_stock * 0.9  # 补到90%容量
                order = max(0, target - inventory)
                # 考虑运输提前期，提前2个月下单
                delivery_month = min(t + self.transport_time, months - 1)
                if delivery_month < months:
                    orders[delivery_month] += order
            
            # 当期到货（考虑之前下单的）
            if t >= self.transport_time:
                actual_arrival = orders[t - self.transport_time]
            else:
                actual_arrival = 0
            
            inventory += actual_arrival
            
            # 满足需求
            if inventory >= self.monthly_demand[t]:
                inventory -= self.monthly_demand[t]
            else:
                # 缺货！钢厂减产
                shortage = self.monthly_demand[t] - inventory
                total_cost += shortage * self.shortage_penalty
                shortage_count += 1
                inventory = 0
            
            # 库存约束
            inventory = min(inventory, self.max_stock)
            
            # 成本计算
            if actual_arrival > 0:
                # 实际支付的是下单时的价格
                order_month = t - self.transport_time
                if order_month >= 0:
                    total_cost += actual_arrival * prices[order_month]
            
            total_cost += inventory * self.holding_cost
        
        return {
            'strategy': '保产量最大化',
            'total_cost': total_cost,
            'shortage_count': shortage_count,
            'final_inventory': inventory,
            'orders': orders
        }
    
    def strategy_min_inventory(self, prices, price_threshold=None):
        """
        策略二：最低安全库存策略
        核心逻辑：JIT准时采购，只在必要时买，且只在低价时买
        
        决策规则：
        - 库存 < 安全库存 → 触发采购
        - 只有价格 < 移动平均线 * 0.95 时才买
        - 小批量多批次
        """
        months = 12
        orders = np.zeros(months)
        inventory = self.initial_stock
        total_cost = 0.0
        shortage_count = 0
        
        # 计算移动平均线作为参考
        if price_threshold is None:
            price_threshold = np.mean(prices) * 0.95
        
        for t in range(months):
            # 低库存策略：只有跌破安全线且价格合适才买
            if inventory <= self.min_safety_stock:
                if prices[t] <= price_threshold:
                    # 价格合适，补到1.5倍安全线
                    target = 1.5 * self.min_safety_stock
                    order = max(0, target - inventory)
                    orders[t] = order
            
            # 到货（铁矿石2个月船期）
            if t >= self.transport_time:
                actual_arrival = orders[t - self.transport_time]
            else:
                actual_arrival = 0
            
            inventory += actual_arrival
            
            # 满足需求
            if inventory >= self.monthly_demand[t]:
                inventory -= self.monthly_demand[t]
            else:
                shortage = self.monthly_demand[t] - inventory
                total_cost += shortage * self.shortage_penalty
                shortage_count += 1
                inventory = 0
            
            inventory = min(inventory, self.max_stock)
            
            # 成本
            if actual_arrival > 0:
                order_month = t - self.transport_time
                if order_month >= 0:
                    total_cost += actual_arrival * prices[order_month]
            
            total_cost += inventory * self.holding_cost
        
        return {
            'strategy': '最低安全库存',
            'total_cost': total_cost,
            'shortage_count': shortage_count,
            'final_inventory': inventory,
            'orders': orders,
            'price_threshold': price_threshold
        }
    
    def strategy_smart_optimization(self, prices):
        """
        策略三：智能优化策略（简单版）
        结合期货升贴水、季节性、库存水平综合决策
        """
        months = 12
        orders = np.zeros(months)
        inventory = self.initial_stock
        total_cost = 0.0
        shortage_count = 0
        
        # 季节性系数：淡季低价囤货，旺季前提前备货
        seasonal_buy = [1.2, 1.3, 0.8, 0.7, 0.9, 1.0, 
                       1.2, 1.1, 0.9, 0.8, 0.7, 0.8]  # 越高越应该买
        
        for t in range(months):
            # 计算目标库存（动态调整）
            # 基础安全库存 + 季节性调整
            target_inventory = self.min_safety_stock * (1 + 0.3 * seasonal_buy[t])
            
            # 价格因素：低价多囤，高价少囤
            price_ratio = prices[t] / np.mean(prices)
            if price_ratio < 0.9:  # 价格低于平均10%
                target_inventory *= 1.3  # 多囤30%
            elif price_ratio > 1.1:  # 价格高于平均10%
                target_inventory *= 0.8  # 少囤20%
            
            # 决策
            if inventory < target_inventory:
                order = min(target_inventory - inventory, 
                           self.max_stock - inventory)
                # 提前2个月下单
                delivery = min(t + self.transport_time, months - 1)
                if delivery < months:
                    orders[delivery] += order
            
            # 到货
            if t >= self.transport_time:
                actual_arrival = orders[t - self.transport_time]
            else:
                actual_arrival = 0
            
            inventory += actual_arrival
            
            # 满足需求
            if inventory >= self.monthly_demand[t]:
                inventory -= self.monthly_demand[t]
            else:
                shortage = self.monthly_demand[t] - inventory
                total_cost += shortage * self.shortage_penalty
                shortage_count += 1
                inventory = 0
            
            inventory = min(inventory, self.max_stock)
            
            # 成本
            if actual_arrival > 0:
                order_month = t - self.transport_time
                if order_month >= 0:
                    total_cost += actual_arrival * prices[order_month]
            
            total_cost += inventory * self.holding_cost
        
        return {
            'strategy': '智能优化策略',
            'total_cost': total_cost,
            'shortage_count': shortage_count,
            'final_inventory': inventory,
            'orders': orders
        }
    
    def compare_strategies(self, n_scenarios=50):
        """对比三种策略"""
        price_scenarios = self.simulate_prices(n_scenarios=n_scenarios)
        
        results = {
            '保产量最大化': [],
            '最低安全库存': [],
            '智能优化策略': []
        }
        
        for i in range(n_scenarios):
            prices = price_scenarios[i]
            
            r1 = self.strategy_max_production(prices)
            r2 = self.strategy_min_inventory(prices)
            r3 = self.strategy_smart_optimization(prices)
            
            results['保产量最大化'].append(r1['total_cost'])
            results['最低安全库存'].append(r2['total_cost'])
            results['智能优化策略'].append(r3['total_cost'])
        
        return results, price_scenarios


def run_demo():
    """运行演示"""
    print("=" * 70)
    print("铁矿石期货采购优化模型")
    print("=" * 70)
    
    # 创建模型
    model = IronOreProcurementModel(
        annual_demand=12000,      # 年需求1200万吨
        initial_stock=800,        # 初始库存800万吨
        min_safety_stock=300,     # 安全库存300万吨
        max_stock=2000,           # 最大库存2000万吨
        holding_cost=15,          # 库存成本15元/吨/月
        shortage_penalty=200,     # 缺货惩罚200元/吨
        transport_time=2          # 运输周期2个月
    )
    
    print(f"\n【基础参数】")
    print(f"年度需求: {model.annual_demand:,} 万吨")
    print(f"初始库存: {model.initial_stock} 万吨")
    print(f"安全库存: {model.min_safety_stock} 万吨")
    print(f"最大库存: {model.max_stock} 万吨")
    print(f"运输周期: {model.transport_time} 个月（巴西矿）")
    
    print(f"\n【月度需求（考虑季节性）】")
    months = ['1月', '2月', '3月', '4月', '5月', '6月', 
              '7月', '8月', '9月', '10月', '11月', '12月']
    for m, d in zip(months, model.monthly_demand):
        print(f"  {m}: {d:6.1f} 万吨")
    
    # 模拟一组价格情景
    np.random.seed(42)
    prices = model.simulate_prices(n_scenarios=1)[0]
    
    print(f"\n【模拟价格路径】")
    for m, p in zip(months, prices):
        print(f"  {m}: {p:6.2f} 元/吨")
    
    # 三种策略对比
    print("\n" + "=" * 70)
    print("【策略对比】")
    print("=" * 70)
    
    r1 = model.strategy_max_production(prices)
    r2 = model.strategy_min_inventory(prices)
    r3 = model.strategy_smart_optimization(prices)
    
    strategies = [r1, r2, r3]
    
    for r in strategies:
        print(f"\n策略: {r['strategy']}")
        print(f"  年度总成本: {r['total_cost']:>12,.2f} 万元")
        print(f"  缺货次数: {r['shortage_count']:>12} 次")
        print(f"  期末库存: {r['final_inventory']:>12.1f} 万吨")
        
        # 成本构成分析
        # 这里简化计算，实际应该拆解
        print(f"  平均库存水平: {np.mean([model.initial_stock, r['final_inventory']]):.1f} 万吨")
    
    # 蒙特卡洛模拟
    print("\n" + "=" * 70)
    print("【蒙特卡洛模拟评估】(50个价格情景)")
    print("=" * 70)
    
    results, _ = model.compare_strategies(n_scenarios=50)
    
    for name, costs in results.items():
        costs = np.array(costs) / 10000  # 转为亿元
        print(f"\n{name}:")
        print(f"  平均成本: {np.mean(costs):.2f} 亿元")
        print(f"  成本波动: {np.std(costs):.2f} 亿元")
        print(f"  最低成本: {np.min(costs):.2f} 亿元")
        print(f"  最高成本: {np.max(costs):.2f} 亿元")
    
    # 可视化
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    # 价格路径
    ax = axes[0, 0]
    ax.plot(range(1, 13), prices, 'b-o', linewidth=2, markersize=6)
    ax.axhline(np.mean(prices), color='r', linestyle='--', 
               label=f'均价: {np.mean(prices):.0f}元')
    ax.set_xlabel('Month')
    ax.set_ylabel('Price (CNY/ton)')
    ax.set_title('Iron Ore Price Simulation (62%Fe)')
    ax.set_xticks(range(1, 13))
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    # 策略成本对比
    ax = axes[0, 1]
    strategy_names = [r['strategy'] for r in strategies]
    costs = [r['total_cost']/10000 for r in strategies]  # 亿元
    colors = ['#e74c3c', '#3498db', '#2ecc71']
    bars = ax.bar(strategy_names, costs, color=colors, alpha=0.8, edgecolor='black')
    ax.set_ylabel('Total Cost (100M CNY)')
    ax.set_title('Strategy Comparison')
    for bar, cost in zip(bars, costs):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                f'{cost:.1f}', ha='center', va='bottom', fontsize=11)
    ax.grid(True, alpha=0.3, axis='y')
    
    # 库存变化对比
    ax = axes[1, 0]
    # 计算各策略的库存轨迹（简化）
    x = range(1, 13)
    ax.fill_between(x, model.min_safety_stock, model.max_stock, 
                    alpha=0.1, color='gray', label='Safe Zone')
    ax.axhline(model.min_safety_stock, color='r', linestyle='--', 
               label=f'Safety Stock ({model.min_safety_stock})')
    
    # 模拟各策略库存（简化展示）
    inv_max = [model.initial_stock - sum(model.monthly_demand[:i]) + 
               sum(r1['orders'][max(0, i-2):i]) for i in range(12)]
    inv_min = [model.initial_stock - sum(model.monthly_demand[:i]) + 
               sum(r2['orders'][max(0, i-2):i]) for i in range(12)]
    
    ax.plot(x, [max(0, v) for v in inv_max], 'r-s', label='Max Production', alpha=0.7)
    ax.plot(x, [max(0, v) for v in inv_min], 'b-^', label='Min Inventory', alpha=0.7)
    ax.set_xlabel('Month')
    ax.set_ylabel('Inventory (10k tons)')
    ax.set_title('Inventory Level Comparison')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    # 成本分布（蒙特卡洛）
    ax = axes[1, 1]
    for name, costs in results.items():
        ax.hist(np.array(costs)/10000, bins=15, alpha=0.5, label=name, edgecolor='black')
    ax.set_xlabel('Total Cost (100M CNY)')
    ax.set_ylabel('Frequency')
    ax.set_title('Cost Distribution (MC Simulation)')
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig('/root/.openclaw/workspace/iron_ore_optimization.png', dpi=150)
    print(f"\n图表已保存: iron_ore_optimization.png")
    
    return model, results


if __name__ == '__main__':
    model, results = run_demo()
