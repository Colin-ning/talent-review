import json
from zhipuai import ZhipuAI

class AIAnalyzer:
    """智谱AI分析器"""
    
    def __init__(self, api_key):
        self.client = ZhipuAI(api_key=api_key)
        self.model = "glm-4"
    
    def analyze_overall(self, df, stats):
        """整体分析"""
        prompt = f"""
你是一位资深的人力资源专家，请基于以下人才盘点数据进行专业分析：

## 盘点数据概览
- 总人数：{stats['total']}人
- 完全匹配：{stats['full_match']}人（{stats['full_match_rate']:.1f}%）
- 基本匹配：{stats['basic_match']}人（{stats['basic_match_rate']:.1f}%）
- 不匹配：{stats['no_match']}人（{stats['no_match_rate']:.1f}%）

## 人才分层
- 高潜人才：{stats['high_potential']}人（{stats['high_potential_rate']:.1f}%）
- 稳定人员：{stats['stable']}人（{stats['stable_rate']:.1f}%）
- 核心骨干：{stats['core_backbone']}人（{stats['core_backbone_rate']:.1f}%）
- 待关注：{stats['attention']}人（{stats['attention_rate']:.1f}%）
- 待优化：{stats['optimization']}人（{stats['optimization_rate']:.1f}%）
- 离职风险：{stats['risk']}人（{stats['risk_rate']:.1f}%）

## 岗位分类
- A类（核心岗位）：{stats['a_class']}人
- M类（中坚力量）：{stats['m_class']}人
- E类（专业效能）：{stats['e_class']}人

请从以下角度进行分析：
1. 团队整体健康度评估
2. 人才结构合理性分析
3. 潜在风险识别
4. 关键发现与洞察

请用专业、简洁的语言进行分析，每点不超过100字。
"""
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}]
        )
        
        return response.choices[0].message.content
    
    def generate_suggestions(self, df, stats):
        """生成优化建议"""
        prompt = f"""
你是一位资深的人力资源专家，请基于以下人才盘点数据提供具体的优化建议：

## 盘点数据
- 总人数：{stats['total']}人
- 高潜人才：{stats['high_potential']}人
- 核心骨干：{stats['core_backbone']}人
- 待优化：{stats['optimization']}人
- 离职风险：{stats['risk']}人
- 学历门槛未通过：{stats['edu_fail']}人

请针对以下方面提供具体、可操作的建议：
1. 高潜人才培养策略
2. 核心骨干保留措施
3. 待优化人员处理方案
4. 离职风险防控措施
5. 团队梯队建设建议

每条建议请包含：问题描述、具体措施、预期效果。用简洁专业的语言表达。
"""
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}]
        )
        
        return response.choices[0].message.content
    
    def analyze_employee(self, employee_data):
        """分析单个员工"""
        prompt = f"""
你是一位资深的人力资源专家，请对以下员工进行个人发展分析：

## 员工信息
- 姓名：{employee_data.get('姓名', '')}
- 部门：{employee_data.get('部门', '')}
- 岗位：{employee_data.get('岗位', '')}
- 职级：{employee_data.get('职级', '')}
- 岗位分类：{employee_data.get('岗位分类', '')}
- 司龄：{employee_data.get('司龄', 0):.1f}年
- 学历档位：{employee_data.get('学历档位', '')}
- 绩效档位：{employee_data.get('绩效档位', '')}
- 匹配度：{employee_data.get('匹配度', '')}
- 人才分层：{employee_data.get('人才分层', '')}

请从以下角度进行分析：
1. 当前状态评估
2. 优势与不足
3. 发展建议
4. 潜在风险提示

请用专业、客观的语言进行分析，给出具体的建议。
"""
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}]
        )
        
        return response.choices[0].message.content
    
    def chat(self, question, context=""):
        """智能问答"""
        system_prompt = """
你是一位专业的人力资源专家，专注于人才盘点和人才管理领域。
你的职责是：
1. 解答关于人才盘点的专业问题
2. 提供人才管理方面的建议
3. 帮助用户理解人才盘点的方法和标准
4. 基于数据给出专业的分析意见

请用专业、友好、简洁的语言回答问题。如果问题超出人才盘点范围，请礼貌地说明。
"""
        
        messages = [
            {"role": "system", "content": system_prompt},
        ]
        
        if context:
            messages.append({"role": "user", "content": f"当前盘点数据：{context}\n\n问题：{question}"})
        else:
            messages.append({"role": "user", "content": question})
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages
        )
        
        return response.choices[0].message.content
    
    def generate_report_summary(self, df, stats):
        """生成报告摘要"""
        prompt = f"""
请为以下人才盘点报告生成一份适合向管理层汇报的摘要：

## 盘点数据
- 盘点总人数：{stats['total']}人
- 部门覆盖：{stats['departments']}个部门
- 岗位类型：{stats['positions']}种职位

## 核心指标
- 人岗匹配率：{stats['full_match_rate']:.1f}%
- 高潜人才占比：{stats['high_potential_rate']:.1f}%
- 核心骨干占比：{stats['core_backbone_rate']:.1f}%
- 待优化人员占比：{stats['optimization_rate']:.1f}%
- 离职风险占比：{stats['risk_rate']:.1f}%

## 岗位分布
- A类岗位：{stats['a_class']}人
- M类岗位：{stats['m_class']}人
- E类岗位：{stats['e_class']}人

请生成：
1. 一句话核心结论（不超过30字）
2. 三个关键发现（每点不超过50字）
3. 三条核心建议（每点不超过50字）

格式要求：简洁明了，适合PPT汇报。
"""
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}]
        )
        
        return response.choices[0].message.content
