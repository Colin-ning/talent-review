import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

st.set_page_config(
    page_title="产研团队人才盘点系统",
    page_icon="📊",
    layout="wide"
)

if 'df' not in st.session_state:
    st.session_state.df = None
if 'result_df' not in st.session_state:
    st.session_state.result_df = None
if 'stats' not in st.session_state:
    st.session_state.stats = None

def get_education_tier(row):
    school_type = row.get('学历-学校类别', '')
    if pd.isna(school_type):
        school_type = ''
    else:
        school_type = str(school_type)
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

def get_performance_tier(row):
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

def check_min_education(row):
    school_type = row.get('学历-学校类别', '')
    if pd.isna(school_type):
        school_type = ''
    else:
        school_type = str(school_type)
    qs_rank = row.get('海外高校QS排名', None)
    position_class = row.get('岗位分类', '')
    if pd.isna(position_class):
        position_class = ''
    else:
        position_class = str(position_class)
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

def check_match(row):
    if not row.get('学历门槛通过', True):
        return '不匹配'
    edu_tier = row.get('学历档位', '三档')
    perf_tier = row.get('绩效档位', '三档')
    position_class = row.get('岗位分类', '')
    if pd.isna(position_class):
        position_class = ''
    else:
        position_class = str(position_class)
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

def is_high_potential(row):
    performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
    recent_2_perfs = []
    for col in performance_cols[-2:]:
        if col in row.index and pd.notna(row[col]) and row[col] != '-':
            recent_2_perfs.append(row[col])
    if len(recent_2_perfs) < 2:
        return False
    recent_2_all_as = all(p in ['S', 'A'] for p in recent_2_perfs)
    young = row.get('司龄', 0) <= 1
    level = row.get('职级', '')
    if pd.isna(level):
        level = ''
    else:
        level = str(level)
    level_ok = level in ['L3', 'L4', 'L5', 'L6', '培训生']
    return recent_2_all_as and young and level_ok

def is_stable(row):
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

def is_core_backbone(row):
    position_class = row.get('岗位分类', '')
    if pd.isna(position_class):
        position_class = ''
    else:
        position_class = str(position_class)
    perf_tier = row.get('绩效档位', '')
    if pd.isna(perf_tier):
        perf_tier = ''
    else:
        perf_tier = str(perf_tier)
    level = row.get('职级', '')
    if pd.isna(level):
        level = ''
    else:
        level = str(level)
    class_ok = position_class in ['A类', 'M类']
    perf_ok = perf_tier in ['一档', '二档']
    time_ok = row.get('司龄', 0) >= 0.5
    level_ok = level in ['L6', 'L7', 'L8', 'L9', 'M4']
    match_ok = row.get('匹配度', '') == '完全匹配'
    return class_ok and perf_ok and time_ok and level_ok and match_ok

def needs_attention(row):
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
    company_years = row.get('司龄', 0)
    if 2 <= company_years <= 3:
        valid_perfs = [p for p in performances if p != '-']
        if len(valid_perfs) >= 2:
            perf_order = {'S': 5, 'A': 4, 'B+': 3, 'B': 2, 'C': 1}
            if perf_order.get(valid_perfs[-1], 0) < perf_order.get(valid_perfs[-2], 0):
                return True
    return False

def needs_optimization(row):
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

def is_resignation_risk(row):
    company_years = row.get('司龄', 0)
    if 2 <= company_years <= 3:
        performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        performances = []
        for col in performance_cols:
            if col in row.index and pd.notna(row[col]) and row[col] != '-':
                performances.append(row[col])
        if len(performances) >= 2:
            perf_order = {'S': 5, 'A': 4, 'B+': 3, 'B': 2, 'C': 1}
            if perf_order.get(performances[-1], 0) < perf_order.get(performances[-2], 0):
                return True
    if company_years < 0.5:
        performance_cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        performances = []
        for col in performance_cols:
            if col in row.index and pd.notna(row[col]) and row[col] != '-':
                performances.append(row[col])
        if any(p in ['B', 'C'] for p in performances):
            return True
    position_class = row.get('岗位分类', '')
    if pd.isna(position_class):
        position_class = ''
    else:
        position_class = str(position_class)
    if (row.get('待关注', False) or row.get('待优化', False)) and position_class == 'A类':
        return True
    return False

def perform_review(df):
    df = df.copy()
    df['学历档位'] = df.apply(get_education_tier, axis=1)
    df['绩效档位'] = df.apply(get_performance_tier, axis=1)
    df['学历门槛通过'] = df.apply(check_min_education, axis=1)
    df['匹配度'] = df.apply(check_match, axis=1)
    today = datetime.now()
    df['司龄'] = df['入职时间'].apply(lambda x: (today - pd.to_datetime(x)).days / 365 if pd.notna(x) else 0)
    df['高潜人才'] = df.apply(is_high_potential, axis=1)
    df['稳定人员'] = df.apply(is_stable, axis=1)
    df['核心骨干'] = df.apply(is_core_backbone, axis=1)
    df['待关注'] = df.apply(needs_attention, axis=1)
    df['待优化'] = df.apply(needs_optimization, axis=1)
    df['离职风险'] = df.apply(is_resignation_risk, axis=1)
    return df

def calculate_stats(df):
    total = len(df)
    formal = len(df[df['岗位类型'] == '正式']) if '岗位类型' in df.columns else total
    return {
        'total': total,
        'formal': formal,
        'departments': df['部门'].nunique() if '部门' in df.columns else 0,
        'positions': df['职位'].nunique() if '职位' in df.columns else 0,
        'a_class': len(df[df['岗位分类'] == 'A类']),
        'm_class': len(df[df['岗位分类'] == 'M类']),
        'e_class': len(df[df['岗位分类'] == 'E类']),
        'full_match': len(df[df['匹配度'] == '完全匹配']),
        'basic_match': len(df[df['匹配度'] == '基本匹配']),
        'no_match': len(df[df['匹配度'] == '不匹配']),
        'full_match_rate': len(df[df['匹配度'] == '完全匹配']) / total * 100 if total > 0 else 0,
        'basic_match_rate': len(df[df['匹配度'] == '基本匹配']) / total * 100 if total > 0 else 0,
        'no_match_rate': len(df[df['匹配度'] == '不匹配']) / total * 100 if total > 0 else 0,
        'high_potential': int(df['高潜人才'].sum()),
        'stable': int(df['稳定人员'].sum()),
        'core_backbone': int(df['核心骨干'].sum()),
        'attention': int(df['待关注'].sum()),
        'optimization': int(df['待优化'].sum()),
        'risk': int(df['离职风险'].sum()),
        'edu_fail': len(df[~df['学历门槛通过']]),
        'high_potential_rate': df['高潜人才'].sum() / total * 100 if total > 0 else 0,
        'stable_rate': df['稳定人员'].sum() / total * 100 if total > 0 else 0,
        'core_backbone_rate': df['核心骨干'].sum() / total * 100 if total > 0 else 0,
        'attention_rate': df['待关注'].sum() / total * 100 if total > 0 else 0,
        'optimization_rate': df['待优化'].sum() / total * 100 if total > 0 else 0,
        'risk_rate': df['离职风险'].sum() / total * 100 if total > 0 else 0,
    }

def generate_excel_report(df, stats):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        total_employees = stats['total']
        formal_employees = stats['formal']
        departments = stats['departments']
        positions = stats['positions']
        a_class = stats['a_class']
        m_class = stats['m_class']
        e_class = stats['e_class']
        
        report_data = [
            ['产研团队全量人才盘点报告'],
            [f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}'],
            [],
            ['一、盘点概览'],
            [],
            ['本次盘点范围', '产研团队全体员工'],
            ['盘点总人数', f'{total_employees}人（正式员工{formal_employees}人）'],
            ['部门覆盖', f'{departments}个部门'],
            ['岗位类型', f'{positions}种职位'],
            [],
            ['岗位分类分布：'],
            ['A类（核心岗位）', f'{a_class}人', f'{a_class/total_employees*100:.1f}%' if total_employees > 0 else '0%'],
            ['M类（中坚力量）', f'{m_class}人', f'{m_class/total_employees*100:.1f}%' if total_employees > 0 else '0%'],
            ['E类（专业效能）', f'{e_class}人', f'{e_class/total_employees*100:.1f}%' if total_employees > 0 else '0%'],
            [],
            ['核心盘点结论：'],
            ['1', f'团队规模适中，M类岗位占比最高（{m_class/total_employees*100:.1f}%）' if total_employees > 0 else '无数据'],
            ['2', f'职级梯队以L5-L7为主，占比{(len(df[df["职级"].isin(["L5","L6","L7"])])/total_employees*100):.1f}%' if total_employees > 0 else '无数据'],
            ['3', f'人岗匹配度：完全匹配{stats["full_match"]}人({stats["full_match_rate"]:.1f}%)，不匹配{stats["no_match"]}人({stats["no_match_rate"]:.1f}%)'],
        ]
        pd.DataFrame(report_data).to_excel(writer, sheet_name='一、盘点概览', index=False, header=False)
        
        if '部门' in df.columns and '职位' in df.columns:
            dept_position_stats = df.groupby(['部门', '职位']).agg({
                '工号': 'count',
                '职级': lambda x: ', '.join(sorted(x.unique())),
                '岗位分类': 'first'
            }).reset_index()
            dept_position_stats.columns = ['部门', '职位', '人数', '职级分布', '岗位分类']
            dept_position_stats.to_excel(writer, sheet_name='二、部门职位分布', index=False)
        
        level_order = ['L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'M4', '培训生', '外包']
        level_stats_data = []
        for level in level_order:
            level_df = df[df['职级'] == level]
            if len(level_df) > 0:
                depts = ', '.join(level_df['部门'].unique()[:5]) if '部门' in df.columns else ''
                level_stats_data.append([level, len(level_df), f'{len(level_df)/total_employees*100:.1f}%', depts])
        pd.DataFrame(level_stats_data, columns=['职级', '人数', '占比', '主要分布部门']).to_excel(
            writer, sheet_name='三、职级体系分析', index=False)
        
        high_level_low_position = df[(df['职级'].isin(['L8', 'L9', 'M4'])) & (df['岗位分类'] == 'E类')]
        low_level_high_position = df[(df['职级'].isin(['L4', 'L5', '培训生'])) & (df['岗位分类'] == 'A类')]
        mismatch_data = [['职级与岗位不匹配异常清单'], [], ['高职级低配异常（职级高但岗位为E类）']]
        if len(high_level_low_position) > 0:
            for _, row in high_level_low_position.iterrows():
                mismatch_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('职级', ''), 'E类岗位'])
        else:
            mismatch_data.append(['✅ 无高职级低配异常'])
        mismatch_data.append([])
        mismatch_data.append(['低职级高配异常（职级低但岗位为A类）'])
        if len(low_level_high_position) > 0:
            for _, row in low_level_high_position.iterrows():
                mismatch_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('职级', ''), 'A类岗位'])
        else:
            mismatch_data.append(['✅ 无低职级高配异常'])
        pd.DataFrame(mismatch_data).to_excel(writer, sheet_name='四、职级岗位异常', index=False, header=False)
        
        education_check_fail = df[~df['学历门槛通过']]
        education_check_pass = df[df['学历门槛通过']]
        education_data = [['岗位最低学历门槛校验'], [], ['校验结果：'],
            ['通过门槛', f'{len(education_check_pass)}人', f'{len(education_check_pass)/total_employees*100:.1f}%' if total_employees > 0 else '0%'],
            ['未通过门槛', f'{len(education_check_fail)}人', f'{len(education_check_fail)/total_employees*100:.1f}%' if total_employees > 0 else '0%'],
            [], ['不满足最低学历要求人员清单：']]
        for _, row in education_check_fail.iterrows():
            education_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('岗位分类', ''), 
                row.get('学历-学校类别', ''), row.get('学历档位', '')])
        pd.DataFrame(education_data).to_excel(writer, sheet_name='五、学历门槛校验', index=False, header=False)
        
        match_data = [['人岗匹配度校验明细：'], [], 
            ['完全匹配', f'{stats["full_match"]}人', f'{stats["full_match_rate"]:.1f}%'],
            ['基本匹配', f'{stats["basic_match"]}人', f'{stats["basic_match_rate"]:.1f}%'],
            ['不匹配', f'{stats["no_match"]}人', f'{stats["no_match_rate"]:.1f}%'],
            [], ['不符合学历+绩效标准人员清单：']]
        not_match = df[df['匹配度'] == '不匹配']
        for _, row in not_match.iterrows():
            match_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('岗位分类', ''),
                row.get('学历档位', ''), row.get('绩效档位', '')])
        pd.DataFrame(match_data).to_excel(writer, sheet_name='六、人岗匹配度', index=False, header=False)
        
        match_summary = []
        for position_class in ['A类', 'M类', 'E类']:
            class_data = df[df['岗位分类'] == position_class]
            if len(class_data) > 0:
                match_dist = class_data['匹配度'].value_counts()
                for match_type in ['完全匹配', '基本匹配', '不匹配']:
                    count = match_dist.get(match_type, 0)
                    match_summary.append([position_class, match_type, count, f'{count/len(class_data)*100:.1f}%'])
        pd.DataFrame(match_summary, columns=['岗位分类', '匹配度', '人数', '占比']).to_excel(writer, sheet_name='七、匹配度汇总', index=False)
        
        high_potential = df[df['高潜人才']]
        hp_data = [['高潜人才盘点'], [], [f'高潜人才明细（共{len(high_potential)}人，占比{len(high_potential)/total_employees*100:.1f}%）：' if total_employees > 0 else '无数据'],
            [], ['姓名', '部门', '岗位', '职级', '司龄', '岗位分类']]
        for _, row in high_potential.iterrows():
            hp_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('职级', ''), 
                f'{row.get("司龄", 0):.1f}年', row.get('岗位分类', '')])
        if len(high_potential) == 0:
            hp_data.append(['暂无符合条件的高潜人才'])
        pd.DataFrame(hp_data).to_excel(writer, sheet_name='八、高潜人才', index=False, header=False)
        
        stable_staff = df[df['稳定人员']]
        stable_data = [['稳定人员盘点'], [], [f'稳定人员明细（共{len(stable_staff)}人）：'], [], ['姓名', '部门', '岗位', '职级', '司龄', '岗位分类']]
        for _, row in stable_staff.iterrows():
            stable_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('职级', ''),
                f'{row.get("司龄", 0):.1f}年', row.get('岗位分类', '')])
        pd.DataFrame(stable_data).to_excel(writer, sheet_name='九、稳定人员', index=False, header=False)
        
        core_backbone = df[df['核心骨干']]
        cb_data = [['核心骨干人员盘点'], [], [f'核心骨干人员明细（共{len(core_backbone)}人）：'], [], ['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度']]
        for _, row in core_backbone.iterrows():
            cb_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('职级', ''),
                row.get('岗位分类', ''), row.get('匹配度', '')])
        pd.DataFrame(cb_data).to_excel(writer, sheet_name='十、核心骨干', index=False, header=False)
        
        attention_needed = df[df['待关注']]
        at_data = [['待关注人员盘点'], [], [f'待关注人员明细（共{len(attention_needed)}人）：'], [], ['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度', '风险等级']]
        for _, row in attention_needed.iterrows():
            at_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('职级', ''),
                row.get('岗位分类', ''), row.get('匹配度', ''), '中'])
        pd.DataFrame(at_data).to_excel(writer, sheet_name='十一、待关注人员', index=False, header=False)
        
        optimization_needed = df[df['待优化']]
        op_data = [['待优化人员盘点'], [], [f'待优化人员明细（共{len(optimization_needed)}人）：'], [], ['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度', '优化优先级']]
        for _, row in optimization_needed.iterrows():
            op_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('职级', ''),
                row.get('岗位分类', ''), row.get('匹配度', ''), '高'])
        pd.DataFrame(op_data).to_excel(writer, sheet_name='十二、待优化人员', index=False, header=False)
        
        resignation_risk = df[df['离职风险']]
        risk_data = [['离职高风险预警'], [], [f'离职高风险人员清单（共{len(resignation_risk)}人）：'], [], ['姓名', '部门', '岗位', '职级', '岗位分类', '司龄', '风险等级', '干预建议']]
        for _, row in resignation_risk.iterrows():
            risk_level = '高' if row.get('岗位分类', '') == 'A类' else '中'
            risk_data.append([row.get('姓名', ''), row.get('部门', ''), row.get('岗位', ''), row.get('职级', ''),
                row.get('岗位分类', ''), f'{row.get("司龄", 0):.1f}年', risk_level, '定期沟通，了解诉求'])
        pd.DataFrame(risk_data).to_excel(writer, sheet_name='十三、离职高风险', index=False, header=False)
        
        fault_data = [['梯队断层风险预警'], [], ['梯队断层风险清单：']]
        if '职位' in df.columns:
            position_level_check = df.groupby('职位')['职级'].apply(lambda x: any(l in ['L6', 'L7', 'L8', 'L9', 'M4'] for l in x))
            positions_without_senior = position_level_check[~position_level_check].index.tolist()
            if positions_without_senior:
                for position in positions_without_senior:
                    position_data = df[df['职位'] == position]
                    fault_data.append([f'{position}岗位无L6及以上职级人员（共{len(position_data)}人）', '影响：缺乏技术骨干', '建议：优先培养'])
            else:
                fault_data.append(['✅ 各岗位梯队完整，无明显断层风险'])
        else:
            fault_data.append(['✅ 无职位数据'])
        pd.DataFrame(fault_data).to_excel(writer, sheet_name='十四、梯队断层风险', index=False, header=False)
        
        suggestions = [
            f'1. 高潜人才培养：{stats["high_potential"]}人',
            f'2. 核心骨干保留：{stats["core_backbone"]}人',
            f'3. 待优化处理：{stats["optimization"]}人',
            f'4. 离职风险防控：{stats["risk"]}人',
            f'5. 定期盘点机制：建议每半年进行一次人才盘点'
        ]
        suggestion_data = [['核心优化建议'], [], ['基于本次盘点结果，提出以下优化建议：'], []] + [[s] for s in suggestions]
        pd.DataFrame(suggestion_data).to_excel(writer, sheet_name='十五、优化建议', index=False, header=False)
        
        export_cols = ['工号', '姓名', '部门', '岗位', '职位', '职级', '岗位分类', '学历档位', '绩效档位', '匹配度', '司龄', 
                       '高潜人才', '稳定人员', '核心骨干', '待关注', '待优化', '离职风险']
        available_cols = [c for c in export_cols if c in df.columns]
        df[available_cols].to_excel(writer, sheet_name='十六、完整盘点数据', index=False)
    
    output.seek(0)
    return output

def generate_pdf_report(df, stats):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        
        output = BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=22, textColor=colors.HexColor('#2C3E50'), spaceAfter=20, alignment=1)
        heading_style = ParagraphStyle('CustomHeading', parent=styles['Heading2'], fontSize=14, textColor=colors.HexColor('#2C3E50'), spaceAfter=10, spaceBefore=15)
        normal_style = ParagraphStyle('CustomNormal', parent=styles['Normal'], fontSize=10, leading=14)
        
        story = []
        story.append(Paragraph('Talent Review Report', title_style))
        story.append(Paragraph(f'Date: {datetime.now().strftime("%Y-%m-%d")}', normal_style))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('1. Overview', heading_style))
        overview_data = [
            ['Metric', 'Details'],
            ['Total Employees', f'{stats["total"]}'],
            ['Full Match', f'{stats["full_match"]} ({stats["full_match_rate"]:.1f}%)'],
            ['Need Optimization', f'{stats["optimization"]}'],
            ['Risk', f'{stats["risk"]}'],
        ]
        overview_table = Table(overview_data, colWidths=[5*cm, 10*cm])
        overview_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498DB')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ]))
        story.append(overview_table)
        story.append(Spacer(1, 15))
        
        story.append(Paragraph('2. Talent Categories', heading_style))
        talent_data = [
            ['Category', 'Count', 'Percentage'],
            ['High Potential', f'{stats["high_potential"]}', f'{stats["high_potential_rate"]:.1f}%'],
            ['Core Backbone', f'{stats["core_backbone"]}', f'{stats["core_backbone_rate"]:.1f}%'],
            ['Need Optimization', f'{stats["optimization"]}', f'{stats["optimization_rate"]:.1f}%'],
            ['Risk', f'{stats["risk"]}', f'{stats["risk_rate"]:.1f}%'],
        ]
        talent_table = Table(talent_data, colWidths=[5*cm, 5*cm, 5*cm])
        talent_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#8E44AD')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ]))
        story.append(talent_table)
        
        doc.build(story)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"PDF Error: {str(e)}")
        return None

with st.sidebar:
    st.title("📊 人才盘点系统")
    page = st.radio("导航", ["首页", "数据上传", "盘点分析", "报告导出"])

if page == "首页":
    st.header("产研团队人才盘点系统")
    st.markdown("""
    ### 使用步骤
    1. 左侧点击 **数据上传** 上传Excel文件
    2. 点击 **盘点分析** 执行盘点
    3. 点击 **报告导出** 下载报告
    """)

elif page == "数据上传":
    st.header("📤 数据上传")
    uploaded_file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'])
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ 成功读取 {len(df)} 条记录")
            st.dataframe(df.head(10))
            if st.button("确认使用此数据", type="primary"):
                st.session_state.df = df
                st.session_state.result_df = None
                st.session_state.stats = None
                st.success("✅ 数据已加载！")
        except Exception as e:
            st.error(f"❌ 读取失败: {str(e)}")

elif page == "盘点分析":
    st.header("📊 盘点分析")
    if st.session_state.df is None:
        st.warning("⚠️ 请先上传数据")
    else:
        st.markdown(f"当前数据：**{len(st.session_state.df)}** 条记录")
        if st.button("执行盘点分析", type="primary"):
            with st.spinner("分析中..."):
                result_df = perform_review(st.session_state.df)
                stats = calculate_stats(result_df)
                st.session_state.result_df = result_df
                st.session_state.stats = stats
            st.success("✅ 完成！")
            st.rerun()
        
        if st.session_state.result_df is not None:
            stats = st.session_state.stats
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("总人数", stats['total'])
            col2.metric("完全匹配", f"{stats['full_match']}人")
            col3.metric("待优化", f"{stats['optimization']}人")
            col4.metric("离职风险", f"{stats['risk']}人")
            st.dataframe(st.session_state.result_df.head(20))

elif page == "报告导出":
    st.header("📄 报告导出")
    if st.session_state.result_df is None:
        st.warning("⚠️ 请先完成盘点")
    else:
        df = st.session_state.result_df
        stats = st.session_state.stats
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Excel报告")
            if st.button("生成Excel报告", type="primary"):
                excel_output = generate_excel_report(df, stats)
                st.download_button("下载Excel", excel_output, f"人才盘点报告_{datetime.now().strftime('%Y%m%d')}.xlsx")
        
        with col2:
            st.markdown("#### PDF报告")
            if st.button("生成PDF报告", type="primary"):
                pdf_output = generate_pdf_report(df, stats)
                if pdf_output:
                    st.download_button("下载PDF", pdf_output, f"人才盘点报告_{datetime.now().strftime('%Y%m%d')}.pdf", "application/pdf")
