# 多AI协作自动化流水线模板
# 场景：30分钟生成一份行业研究报告
# 参与AI：Kimi(我)、DeepSeek-R1、Gemini、豆包

import json
import time
from typing import Dict, List, Callable
from dataclasses import dataclass
from datetime import datetime

@dataclass
class TaskResult:
    """任务结果封装"""
    step: int
    agent: str
    input_data: str
    output_data: str
    cost_tokens: int
    time_spent: float
    status: str  # success / failed

class AIAgentChain:
    """
    AI代理链 orchestrator
    管理多个AI的顺序执行和数据流转
    """
    
    def __init__(self):
        self.chain = []
        self.results: List[TaskResult] = []
        
    def add_step(self, name: str, agent: str, prompt_template: str, 
                 condition: Callable = None):
        """
        添加流水线步骤
        
        Args:
            name: 步骤名称
            agent: 使用的AI (kimi/deepseek/gemini/doubao)
            prompt_template: Prompt模板，可用{input}占位符
            condition: 执行条件函数，返回True才执行
        """
        self.chain.append({
            'name': name,
            'agent': agent,
            'template': prompt_template,
            'condition': condition or (lambda x: True)
        })
        return self
    
    def execute(self, initial_input: str, context: Dict = None) -> Dict:
        """
        执行完整流水线
        
        Args:
            initial_input: 初始输入（用户原始需求）
            context: 全局上下文（如行业、时间范围等）
            
        Returns:
            最终输出结果
        """
        current_data = initial_input
        full_context = context or {}
        
        print(f"🚀 开始执行AI流水线 - {datetime.now().strftime('%H:%M:%S')}")
        print(f"📋 初始任务：{initial_input[:100]}...\n")
        
        for i, step in enumerate(self.chain, 1):
            print(f"\n{'='*60}")
            print(f"步骤 {i}/{len(self.chain)}: {step['name']}")
            print(f"执行AI: {step['agent'].upper()}")
            print(f"{'='*60}")
            
            # 检查执行条件
            if not step['condition'](current_data):
                print(f"⏭️  跳过此步骤（条件不满足）")
                continue
            
            start_time = time.time()
            
            try:
                # 构建Prompt
                prompt = step['template'].format(
                    input=current_data,
                    **full_context
                )
                
                # 调用对应的AI（模拟，实际需接入API）
                result = self._call_ai(step['agent'], prompt)
                
                time_spent = time.time() - start_time
                
                # 记录结果
                task_result = TaskResult(
                    step=i,
                    agent=step['agent'],
                    input_data=prompt[:200],
                    output_data=result[:500],
                    cost_tokens=len(prompt) + len(result),
                    time_spent=time_spent,
                    status='success'
                )
                self.results.append(task_result)
                
                # 更新数据流
                current_data = result
                
                print(f"✅ 完成 - 耗时: {time_spent:.2f}s, Token: {task_result.cost_tokens}")
                print(f"📄 输出预览: {result[:200]}...")
                
            except Exception as e:
                print(f"❌ 失败: {str(e)}")
                self.results.append(TaskResult(
                    step=i,
                    agent=step['agent'],
                    input_data=prompt[:200],
                    output_data=str(e),
                    cost_tokens=0,
                    time_spent=time.time() - start_time,
                    status='failed'
                ))
                break
        
        # 生成执行报告
        return self._generate_report(current_data)
    
    def _call_ai(self, agent: str, prompt: str) -> str:
        """
        调用具体AI（模拟实现，实际需接入真实API）
        """
        # 这里应该接入真实的API调用
        # 例如：OpenAI、Kimi、DeepSeek等
        
        agent_configs = {
            'kimi': {
                'model': 'kimi-coding/k2p5',
                'strength': '中文理解、结构化输出',
                'cost_per_1k': 0.006
            },
            'deepseek': {
                'model': 'deepseek-reasoner',
                'strength': '逻辑推理、数学计算',
                'cost_per_1k': 0.033
            },
            'gemini': {
                'model': 'gemini-1.5-pro',
                'strength': '信息检索、多语言',
                'cost_per_1k': 0.32
            },
            'doubao': {
                'model': 'doubao-pro',
                'strength': '内容创作、图像生成',
                'cost_per_1k': 0.10
            }
        }
        
        config = agent_configs.get(agent, {})
        print(f"   🤖 调用 {config.get('model', agent)}")
        print(f"   💪 专长: {config.get('strength', '通用')}")
        
        # 模拟AI响应（实际应调用真实API）
        return f"[{agent.upper()}处理结果] 基于输入'{prompt[:50]}...'的分析完成"
    
    def _generate_report(self, final_output: str) -> Dict:
        """生成执行报告"""
        total_time = sum(r.time_spent for r in self.results)
        total_tokens = sum(r.cost_tokens for r in self.results)
        
        return {
            'final_output': final_output,
            'execution_summary': {
                'total_steps': len(self.chain),
                'completed_steps': len(self.results),
                'total_time': f"{total_time:.2f}s",
                'total_tokens': total_tokens,
                'success_rate': f"{len([r for r in self.results if r.status=='success'])/len(self.results)*100:.1f}%"
            },
            'step_details': [
                {
                    'step': r.step,
                    'agent': r.agent,
                    'time': f"{r.time_spent:.2f}s",
                    'tokens': r.cost_tokens,
                    'status': r.status
                }
                for r in self.results
            ]
        }


# ============================================================
# 实战模板1：行业研究报告生成流水线
# ============================================================

def create_industry_report_chain() -> AIAgentChain:
    """
    创建行业研究报告生成流水线
    
    流程：
    1. Gemini - 信息检索与数据收集
    2. DeepSeek - 数据分析与洞察提炼
    3. Kimi - 报告撰写与结构化
    4. Doubao - 视觉素材生成（可选）
    """
    
    chain = AIAgentChain()
    
    # Step 1: 信息检索 (Gemini)
    chain.add_step(
        name="信息检索与数据收集",
        agent="gemini",
        prompt_template="""
你是一个专业的行业研究助理。请检索并整理关于"{input}"的最新信息。

要求：
1. 收集2024-2025年的最新市场数据
2. 列出主要玩家及其市场份额
3. 找出3-5个关键趋势
4. 格式：JSON格式，包含 market_size, key_players, trends, data_sources

主题：{input}
"""
    )
    
    # Step 2: 数据分析 (DeepSeek-R1)
    chain.add_step(
        name="数据分析与洞察提炼",
        agent="deepseek",
        prompt_template="""
基于以下行业数据，进行深入分析：

{input}

要求：
1. 计算年复合增长率(CAGR)
2. 分析竞争格局（波特五力框架）
3. 识别3个关键洞察（Insight）
4. 评估2个主要风险

输出格式：
- 定量分析：（数据计算）
- 定性分析：（战略洞察）
- 风险提示：
"""
    )
    
    # Step 3: 报告撰写 (Kimi)
    chain.add_step(
        name="报告撰写与结构化",
        agent="kimi",
        prompt_template="""
基于以下分析结果，撰写一份专业的行业研究报告：

{input}

报告要求：
1. 标题：【行业研究】XXX市场深度分析
2. 结构：
   - Executive Summary（执行摘要）
   - 市场规模与增长趋势
   - 竞争格局分析
   - 关键洞察与建议
   - 风险提示
3. 风格：专业、客观、适合向高管汇报
4. 字数：2000-3000字

请直接输出完整的Markdown格式报告。
"""
    )
    
    # Step 4: 视觉素材生成 (Doubao) - 条件执行
    chain.add_step(
        name="生成报告配图",
        agent="doubao",
        prompt_template="""
为以下行业研究报告生成一张高质量的封面图：

报告主题：{input}

要求：
- 风格：科技感、商务风
- 配色：蓝色系（专业感）
- 元素：包含行业相关的视觉符号
- 尺寸：16:9 横版

Prompt（用于AI绘图）:
"""
    )
    
    return chain


# ============================================================
# 实战模板2：代码开发流水线
# ============================================================

def create_code_development_chain() -> AIAgentChain:
    """
    代码开发协作流水线
    
    流程：
    1. Kimi - 需求分析与架构设计
    2. DeepSeek - 核心算法实现
    3. Kimi - 代码优化与文档
    4. Gemini - 代码审查（可选）
    """
    
    chain = AIAgentChain()
    
    # Step 1: 需求分析
    chain.add_step(
        name="需求分析与架构设计",
        agent="kimi",
        prompt_template="""
分析以下编程需求，并设计解决方案：

需求：{input}

输出：
1. 功能清单（用✓标记核心功能）
2. 技术架构图（用文字描述）
3. 关键类/函数设计
4. 依赖库建议
"""
    )
    
    # Step 2: 核心实现
    chain.add_step(
        name="核心算法实现",
        agent="deepseek",
        prompt_template="""
基于以下设计，实现核心代码：

{input}

要求：
1. 使用Python 3.9+
2. 包含类型提示(type hints)
3. 关键逻辑添加注释
4. 处理边界情况和异常
5. 提供2-3个使用示例

只输出代码，不需要解释。
"""
    )
    
    # Step 3: 优化与文档
    chain.add_step(
        name="代码优化与文档",
        agent="kimi",
        prompt_template="""
优化以下代码，并添加完整文档：

{input}

要求：
1. 添加Google风格的docstring
2. 优化变量命名和代码结构
3. 添加单元测试示例
4. 创建README.md（包含安装、使用、API说明）
5. 检查PEP8规范
"""
    )
    
    return chain


# ============================================================
# 实战模板3：内容创作流水线
# ============================================================

def create_content_creation_chain() -> AIAgentChain:
    """
    多平台内容创作流水线
    
    流程：
    1. Gemini - 热点趋势分析
    2. DeepSeek - 内容逻辑梳理
    3. Kimi - 文案撰写
    4. Doubao - 配图生成
    """
    
    chain = AIAgentChain()
    
    # Step 1: 趋势分析
    chain.add_step(
        name="热点趋势分析",
        agent="gemini",
        prompt_template="""
分析当前关于"{input}"的社交媒体趋势：

1. 找出3个热门话题
2. 分析目标受众画像
3. 列出5个高频关键词
4. 建议最佳发布时间

输出JSON格式。
"""
    )
    
    # Step 2: 内容逻辑
    chain.add_step(
        name="内容逻辑梳理",
        agent="deepseek",
        prompt_template="""
基于以下趋势数据，设计内容逻辑框架：

{input}

要求：
1. 设计一个吸引人的开场（Hook）
2. 构建3个内容要点（逻辑递进）
3. 设计互动环节（CTA）
4. 控制字数：小红书300字/知乎1500字/公众号2000字
"""
    )
    
    # Step 3: 文案撰写
    chain.add_step(
        name="多平台文案撰写",
        agent="kimi",
        prompt_template="""
基于以下内容框架，撰写3个版本的文案：

{input}

输出：
1. 小红书版（emoji多、口语化、带话题标签）
2. 知乎版（专业、深度、结构化）
3. 公众号版（故事性、可读性强）

每个版本包含：标题、正文、标签/话题
"""
    )
    
    # Step 4: 配图生成
    chain.add_step(
        name="生成配图Prompt",
        agent="doubao",
        prompt_template="""
为以下内容生成3张配图的AI绘图Prompt：

{input}

要求：
1. 风格统一（国潮/科技/商务）
2. 符合各平台调性
3. 包含具体视觉元素描述
4. 指定色彩方案

输出可直接复制到Midjourney/Stable Diffusion的Prompt。
"""
    )
    
    return chain


# ============================================================
# 使用示例
# ============================================================

def demo():
    """演示如何使用流水线"""
    
    print("="*70)
    print("多AI协作自动化流水线 - 演示")
    print("="*70)
    
    # 选择流水线类型
    print("\n选择流水线：")
    print("1. 行业研究报告")
    print("2. 代码开发")
    print("3. 内容创作")
    
    # 示例：行业研究报告
    chain = create_industry_report_chain()
    
    # 执行任务
    result = chain.execute(
        initial_input="中国工业大模型市场现状与趋势",
        context={
            'industry': 'AI/工业',
            'time_range': '2024-2025',
            'language': '中文'
        }
    )
    
    # 输出结果
    print("\n" + "="*70)
    print("执行报告")
    print("="*70)
    print(json.dumps(result['execution_summary'], indent=2, ensure_ascii=False))
    
    print("\n" + "="*70)
    print("最终输出预览")
    print("="*70)
    print(result['final_output'][:500] + "...")
    
    return result


if __name__ == '__main__':
    result = demo()
