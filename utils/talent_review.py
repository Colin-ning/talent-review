import pandas as pd
import numpy as np
from datetime import datetime

class TalentReviewer:
    """人才盘点核心逻辑类"""
    
    def __init__(self, df):
        self.df = df.copy()
        self.result_df = None
        self.stats = None
    
    def get_education_tier(self, row):
        """判定学历档位"""
        school_type = row.get('学历-学校类别', '')
        qs_rank = row.get('海外高校QS排名', None)
        
        tier1_schools = ['C9', '985']
        tier2_schools = ['211', '双一流', '类211']
        
        if pd.notna(qs_rank) and qs_rank <= 300:
            return '一档'
        
        if school_type in tier1_schools:
            return '一档'
        
        if pd.notna(qs_rank) and qs_rank <= 500:
            return '二档'
        
        if school_type in tier2_schools:
            return '二档'
        
        if school_type in ['一本', '二本', '三本', '大专']:
            return '三档'
        
        if school_type == '海外高校':
            if pd.notna(qs_rank):
                if qs_rank <= 300:
                    return '一档'
                elif qs_rank <= 500:
                    return '二档'
                else:
                    return '三档'
            return '三档'
        
        return '三档'
    
    def get_performance_tier(self, row):
        """判定绩效档位"""
        performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        performances = []
        for col in performance_cols:
            if col in row.index and pd.notna(row[col]) and row[col] != '-':
                performances.append(row[col])
        
        if len(performances) == 0:
            return '二档'
        
        s_a_count = sum(1 for p in performances if p in ['S', 'A'])
        total_count = len(performances)
        
        has_annual_sa = False
        annual_cols = ['2024年度绩效结果', '2025年度绩效结果']
        for col in annual_cols:
            if col in row.index and pd.notna(row[col]) and row[col] in ['S', 'A']:
                has_annual_sa = True
                break
        
        all_b_plus_or_above = all(p in ['S', 'A', 'B+'] for p in performances)
        b_count = sum(1 for p in performances if p == 'B')
        
        if s_a_count / total_count > 0.5 and has_annual_sa and all_b_plus_or_above:
            return '一档'
        
        if all_b_plus_or_above:
            return '二档'
        
        if b_count <= 1 and all(p in ['S', 'A', 'B+', 'B'] for p in performances):
            return '三档'
        
        return '三档'
    
    def check_min_education(self, row):
        """检查学历门槛"""
        school_type = row.get('学历-学校类别', '')
        qs_rank = row.get('海外高校QS排名', None)
        position_class = row.get('岗位分类', '')
        
        tier1_schools = ['C9', '985']
        tier2_schools = ['211', '双一流', '类211']
        
        if position_class == 'A类':
            if school_type in tier1_schools or school_type in tier2_schools:
                return True
            if pd.notna(qs_rank) and qs_rank <= 500:
                return True
            return False
        
        elif position_class == 'M类':
            if school_type in tier1_schools or school_type in tier2_schools or school_type == '一本':
                return True
            if pd.notna(qs_rank) and qs_rank <= 1000:
                return True
            return False
        
        elif position_class == 'E类':
            if school_type in tier1_schools or school_type in tier2_schools or school_type in ['一本', '二本']:
                return True
            return False
        
        return True
    
    def check_match(self, row):
        """判定人岗匹配度"""
        if not row.get('学历门槛通过', True):
            return '不匹配'
        
        edu_tier = row.get('学历档位', '三档')
        perf_tier = row.get('绩效档位', '三档')
        position_class = row.get('岗位分类', '')
        
        performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        performances = []
        for col in performance_cols:
            if col in row.index and pd.notna(row[col]) and row[col] != '-':
                performances.append(row[col])
        
        has_c = any(p == 'C' for p in performances)
        
        if has_c:
            return '不匹配'
        
        if position_class == 'A类':
            if (edu_tier == '一档' and perf_tier == '一档') or \
               (edu_tier == '二档' and perf_tier == '一档') or \
               (edu_tier == '一档' and perf_tier == '二档'):
                return '完全匹配'
        
        elif position_class == 'M类':
            if (edu_tier == '二档' and perf_tier == '二档') or \
               (edu_tier == '一档' and perf_tier == '三档') or \
               (edu_tier == '三档' and perf_tier == '一档'):
                return '完全匹配'
        
        elif position_class == 'E类':
            if (edu_tier == '二档' and perf_tier == '三档') or \
               (edu_tier == '三档' and perf_tier == '二档'):
                return '完全匹配'
        
        return '基本匹配'
    
    def is_high_potential(self, row):
        """判定高潜人才"""
        performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        recent_2_perfs = []
        for col in performance_cols[-2:]:
            if col in row.index and pd.notna(row[col]) and row[col] != '-':
                recent_2_perfs.append(row[col])
        
        if len(recent_2_perfs) < 2:
            return False
        
        recent_2_all_as = all(p in ['S', 'A'] for p in recent_2_perfs)
        young = row.get('司龄', 0) <= 1
        level_ok = row.get('职级', '') in ['L3', 'L4', 'L5', 'L6', '培训生']
        
        return recent_2_all_as and young and level_ok
    
    def is_stable(self, row):
        """判定稳定人员"""
        performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        performances = []
        for col in performance_cols:
            if col in row.index and pd.notna(row[col]) and row[col] != '-':
                performances.append(row[col])
        
        if len(performances) < 4:
            return False
        
        all_b_plus = all(p in ['S', 'A', 'B+'] for p in performances)
        stable_time = row.get('司龄', 0) >= 0.5
        
        return all_b_plus and stable_time
    
    def is_core_backbone(self, row):
        """判定核心骨干"""
        class_ok = row.get('岗位分类', '') in ['A类', 'M类']
        perf_ok = row.get('绩效档位', '') in ['一档', '二档']
        time_ok = row.get('司龄', 0) >= 0.5
        level_ok = row.get('职级', '') in ['L6', 'L7', 'L8', 'L9', 'M4']
        match_ok = row.get('匹配度', '') == '完全匹配'
        
        return class_ok and perf_ok and time_ok and level_ok and match_ok
    
    def needs_attention(self, row):
        """判定待关注人员"""
        performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        performances = []
        for col in performance_cols:
            if col in row.index and pd.notna(row[col]) and row[col] != '-':
                performances.append(row[col])
        
        b_count = sum(1 for p in performances if p == 'B')
        if b_count >= 2:
            return True
        
        if row.get('匹配度', '') == '基本匹配':
            return True
        
        if 2 <= row.get('司龄', 0) <= 3:
            valid_perfs = [p for p in performances if p != '-']
            if len(valid_perfs) >= 2:
                perf_order = {'S': 5, 'A': 4, 'B+': 3, 'B': 2, 'C': 1}
                if perf_order.get(valid_perfs[-1], 0) < perf_order.get(valid_perfs[-2], 0):
                    return True
        
        return False
    
    def needs_optimization(self, row):
        """判定待优化人员"""
        performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        performances = []
        for col in performance_cols:
            if col in row.index and pd.notna(row[col]) and row[col] != '-':
                performances.append(row[col])
        
        if any(p == 'C' for p in performances):
            return True
        
        if row.get('匹配度', '') == '不匹配':
            return True
        
        valid_perfs = [p for p in performances if p != '-']
        if len(valid_perfs) >= 2:
            perf_order = {'S': 5, 'A': 4, 'B+': 3, 'B': 2, 'C': 1}
            if perf_order.get(valid_perfs[-1], 0) < perf_order.get(valid_perfs[-2], 0):
                if valid_perfs[-1] in ['B', 'C']:
                    return True
        
        return False
    
    def is_resignation_risk(self, row):
        """判定离职风险"""
        if 2 <= row.get('司龄', 0) <= 3:
            performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
            performances = []
            for col in performance_cols:
                if col in row.index and pd.notna(row[col]) and row[col] != '-':
                    performances.append(row[col])
            if len(performances) >= 2:
                perf_order = {'S': 5, 'A': 4, 'B+': 3, 'B': 2, 'C': 1}
                if perf_order.get(performances[-1], 0) < perf_order.get(performances[-2], 0):
                    return True
        
        if row.get('司龄', 0) < 0.5:
            performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
            performances = []
            for col in performance_cols:
                if col in row.index and pd.notna(row[col]) and row[col] != '-':
                    performances.append(row[col])
            if any(p in ['B', 'C'] for p in performances):
                return True
        
        if (row.get('待关注', False) or row.get('待优化', False)) and row.get('岗位分类', '') == 'A类':
            return True
        
        return False
    
    def perform_review(self):
        """执行盘点分析"""
        self.df['学历档位'] = self.df.apply(self.get_education_tier, axis=1)
        self.df['绩效档位'] = self.df.apply(self.get_performance_tier, axis=1)
        self.df['学历门槛通过'] = self.df.apply(self.check_min_education, axis=1)
        self.df['匹配度'] = self.df.apply(self.check_match, axis=1)
        
        today = datetime.now()
        self.df['司龄'] = self.df['入职时间'].apply(
            lambda x: (today - pd.to_datetime(x)).days / 365 if pd.notna(x) else 0
        )
        
        self.df['高潜人才'] = self.df.apply(self.is_high_potential, axis=1)
        self.df['稳定人员'] = self.df.apply(self.is_stable, axis=1)
        self.df['核心骨干'] = self.df.apply(self.is_core_backbone, axis=1)
        self.df['待关注'] = self.df.apply(self.needs_attention, axis=1)
        self.df['待优化'] = self.df.apply(self.needs_optimization, axis=1)
        self.df['离职风险'] = self.df.apply(self.is_resignation_risk, axis=1)
        
        def get_talent_layer(row):
            if row['核心骨干']:
                return '核心骨干'
            elif row['高潜人才']:
                return '高潜人才'
            elif row['稳定人员']:
                return '稳定人员'
            elif row['待优化']:
                return '待优化'
            elif row['待关注']:
                return '待关注'
            else:
                return '其他'
        
        self.df['人才分层'] = self.df.apply(get_talent_layer, axis=1)
        
        self.result_df = self.df
        self._calculate_statistics()
    
    def _calculate_statistics(self):
        """计算统计数据"""
        total = len(self.df)
        
        self.stats = {
            'total': total,
            'formal': len(self.df[self.df.get('岗位类型', '') == '正式']) if '岗位类型' in self.df.columns else total,
            'departments': self.df['部门'].nunique() if '部门' in self.df.columns else 0,
            'positions': self.df['职位'].nunique() if '职位' in self.df.columns else 0,
            
            'a_class': len(self.df[self.df['岗位分类'] == 'A类']),
            'm_class': len(self.df[self.df['岗位分类'] == 'M类']),
            'e_class': len(self.df[self.df['岗位分类'] == 'E类']),
            
            'full_match': len(self.df[self.df['匹配度'] == '完全匹配']),
            'basic_match': len(self.df[self.df['匹配度'] == '基本匹配']),
            'no_match': len(self.df[self.df['匹配度'] == '不匹配']),
            
            'full_match_rate': len(self.df[self.df['匹配度'] == '完全匹配']) / total * 100,
            'basic_match_rate': len(self.df[self.df['匹配度'] == '基本匹配']) / total * 100,
            'no_match_rate': len(self.df[self.df['匹配度'] == '不匹配']) / total * 100,
            
            'high_potential': int(self.df['高潜人才'].sum()),
            'stable': int(self.df['稳定人员'].sum()),
            'core_backbone': int(self.df['核心骨干'].sum()),
            'attention': int(self.df['待关注'].sum()),
            'optimization': int(self.df['待优化'].sum()),
            'risk': int(self.df['离职风险'].sum()),
            
            'high_potential_rate': self.df['高潜人才'].sum() / total * 100,
            'stable_rate': self.df['稳定人员'].sum() / total * 100,
            'core_backbone_rate': self.df['核心骨干'].sum() / total * 100,
            'attention_rate': self.df['待关注'].sum() / total * 100,
            'optimization_rate': self.df['待优化'].sum() / total * 100,
            'risk_rate': self.df['离职风险'].sum() / total * 100,
            
            'edu_fail': len(self.df[~self.df['学历门槛通过']]),
        }
    
    def get_result_df(self):
        """获取结果DataFrame"""
        return self.result_df
    
    def get_statistics(self):
        """获取统计数据"""
        return self.stats
