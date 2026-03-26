import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(
    page_title="产研团队人才盘点系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

if 'df' not in st.session_state:
    st.session_state.df = None
if 'result_df' not in st.session_state:
    st.session_state.result_df = None
if 'stats' not in st.session_state:
    st.session_state.stats = None

def get_education_tier(row):
    school_type = str(row.get('学历-学校类别', ''))
    qs_rank = row.get('海外高校QS排名', None)
    tier1_schools = ['C9', '985']
    tier2_schools = ['211', '双一流', '类211']
    if pd.notna(qs_rank) and qs_rank <= 300: return '一档'
    if school_type in tier1_schools: return '一档'
    if pd.notna(qs_rank) and qs_rank <= 500: return '二档'
    if school_type in tier2_schools: return '二档'
    if school_type in ['一本', '二本', '三本', '大专']: return '三档'
    if school_type == '海外高校':
        if pd.notna(qs_rank):
            if qs_rank <= 300: return '一档'
            elif qs_rank <= 500: return '二档'
        return '三档'
    return '三档'

def get_performance_tier(row):
    cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
    perfs = [row[c] for c in cols if c in row.index and pd.notna(row[c]) and row[c] != '-']
    if len(perfs) == 0: return '二档'
    sa_count = sum(1 for p in perfs if p in ['S', 'A'])
    has_annual_sa = any(row.get(c) in ['S', 'A'] for c in ['2024年度绩效结果', '2025年度绩效结果'] if c in row.index and pd.notna(row.get(c)))
    all_bp = all(p in ['S', 'A', 'B+'] for p in perfs)
    b_count = sum(1 for p in perfs if p == 'B')
    if sa_count / len(perfs) > 0.5 and has_annual_sa and all_bp: return '一档'
    if all_bp: return '二档'
    if b_count <= 1 and all(p in ['S', 'A', 'B+', 'B'] for p in perfs): return '三档'
    return '三档'

def check_min_education(row):
    school_type = str(row.get('学历-学校类别', ''))
    qs_rank = row.get('海外高校QS排名', None)
    pc = str(row.get('岗位分类', ''))
    tier1, tier2 = ['C9', '985'], ['211', '双一流', '类211']
    if pc == 'A类': return school_type in tier1 + tier2 or (pd.notna(qs_rank) and qs_rank <= 500)
    if pc == 'M类': return school_type in tier1 + tier2 + ['一本'] or (pd.notna(qs_rank) and qs_rank <= 1000)
    if pc == 'E类': return school_type in tier1 + tier2 + ['一本', '二本']
    return True

def check_match(row):
    if not row.get('学历门槛通过', True): return '不匹配'
    edu = row.get('学历档位', '三档')
    perf = row.get('绩效档位', '三档')
    pc = str(row.get('岗位分类', ''))
    cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
    perfs = [row[c] for c in cols if c in row.index and pd.notna(row[c]) and row[c] != '-']
    if any(p == 'C' for p in perfs): return '不匹配'
    if pc == 'A类' and ((edu == '一档' and perf == '一档') or (edu == '二档' and perf == '一档') or (edu == '一档' and perf == '二档')): return '完全匹配'
    if pc == 'M类' and ((edu == '二档' and perf == '二档') or (edu == '一档' and perf == '三档') or (edu == '三档' and perf == '一档')): return '完全匹配'
    if pc == 'E类' and ((edu == '二档' and perf == '三档') or (edu == '三档' and perf == '二档')): return '完全匹配'
    return '基本匹配'

def is_high_potential(row):
    cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
    recent = [row[c] for c in cols[-2:] if c in row.index and pd.notna(row[c]) and row[c] != '-']
    if len(recent) < 2: return False
    return all(p in ['S', 'A'] for p in recent) and row.get('司龄', 0) <= 1 and row.get('职级', '') in ['L3', 'L4', 'L5', 'L6', '培训生']

def is_stable(row):
    cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
    perfs = [row[c] for c in cols if c in row.index and pd.notna(row[c]) and row[c] != '-']
    if len(perfs) < 4: return False
    return all(p in ['S', 'A', 'B+'] for p in perfs) and row.get('司龄', 0) >= 0.5

def is_core_backbone(row):
    return (row.get('岗位分类', '') in ['A类', 'M类'] and 
            row.get('绩效档位', '') in ['一档', '二档'] and 
            row.get('司龄', 0) >= 0.5 and 
            row.get('职级', '') in ['L6', 'L7', 'L8', 'L9', 'M4'] and 
            row.get('匹配度', '') == '完全匹配')

def needs_attention(row):
    cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
    perfs = [row[c] for c in cols if c in row.index and pd.notna(row[c]) and row[c] != '-']
    if sum(1 for p in perfs if p == 'B') >= 2: return True
    if row.get('匹配度', '') == '基本匹配': return True
    if 2 <= row.get('司龄', 0) <= 3 and len(perfs) >= 2:
        order = {'S': 5, 'A': 4, 'B+': 3, 'B': 2, 'C': 1}
        if order.get(perfs[-1], 0) < order.get(perfs[-2], 0): return True
    return False

def needs_optimization(row):
    cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
    perfs = [row[c] for c in cols if c in row.index and pd.notna(row[c]) and row[c] != '-']
    if any(p == 'C' for p in perfs): return True
    if row.get('匹配度', '') == '不匹配': return True
    if len(perfs) >= 2:
        order = {'S': 5, 'A': 4, 'B+': 3, 'B': 2, 'C': 1}
        if order.get(perfs[-1], 0) < order.get(perfs[-2], 0) and perfs[-1] in ['B', 'C']: return True
    return False

def is_resignation_risk(row):
    if 2 <= row.get('司龄', 0) <= 3:
        cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        perfs = [row[c] for c in cols if c in row.index and pd.notna(row[c]) and row[c] != '-']
        if len(perfs) >= 2:
            order = {'S': 5, 'A': 4, 'B+': 3, 'B': 2, 'C': 1}
            if order.get(perfs[-1], 0) < order.get(perfs[-2], 0): return True
    if row.get('司龄', 0) < 0.5:
        cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
        perfs = [row[c] for c in cols if c in row.index and pd.notna(row[c]) and row[c] != '-']
        if any(p in ['B', 'C'] for p in perfs): return True
    if (row.get('待关注', False) or row.get('待优化', False)) and row.get('岗位分类', '') == 'A类': return True
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
    def get_layer(r):
        if r['核心骨干']: return '核心骨干'
        if r['高潜人才']: return '高潜人才'
        if r['稳定人员']: return '稳定人员'
        if r['待优化']: return '待优化人员'
        if r['待关注']: return '待关注人员'
        return '其他人员'
    df['人才分层'] = df.apply(get_layer, axis=1)
    return df

def calculate_stats(df):
    total = len(df)
    return {
        'total': total,
        'formal': len(df[df.get('岗位类型', '') == '正式']) if '岗位类型' in df.columns else total,
        'departments': df['部门'].nunique() if '部门' in df.columns else 0,
        'positions': df['职位'].nunique() if '职位' in df.columns else 0,
        'a_class': len(df[df['岗位分类'] == 'A类']),
        'm_class': len(df[df['岗位分类'] == 'M类']),
        'e_class': len(df[df['岗位分类'] == 'E类']),
        'full_match': len(df[df['匹配度'] == '完全匹配']),
        'basic_match': len(df[df['匹配度'] == '基本匹配']),
        'no_match': len(df[df['匹配度'] == '不匹配']),
        'full_match_rate': len(df[df['匹配度'] == '完全匹配']) / total * 100,
        'basic_match_rate': len(df[df['匹配度'] == '基本匹配']) / total * 100,
        'no_match_rate': len(df[df['匹配度'] == '不匹配']) / total * 100,
        'high_potential': int(df['高潜人才'].sum()),
        'stable': int(df['稳定人员'].sum()),
        'core_backbone': int(df['核心骨干'].sum()),
        'attention': int(df['待关注'].sum()),
        'optimization': int(df['待优化'].sum()),
        'risk': int(df['离职风险'].sum()),
        'edu_fail': len(df[~df['学历门槛通过']]),
        'high_potential_rate': df['高潜人才'].sum() / total * 100,
        'stable_rate': df['稳定人员'].sum() / total * 100,
        'core_backbone_rate': df['核心骨干'].sum() / total * 100,
        'attention_rate': df['待关注'].sum() / total * 100,
        'optimization_rate': df['待优化'].sum() / total * 100,
        'risk_rate': df['离职风险'].sum() / total * 100,
    }

def generate_excel_report(df, stats):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        overview = pd.DataFrame([
            ['产研团队全量人才盘点报告'],
            [f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}'],
            [],
            ['一、盘点概览'],
            [],
            ['本次盘点范围', '产研团队全体员工'],
            ['盘点总人数', f'{stats["total"]}人（正式员工{stats["formal"]}人）'],
            ['部门覆盖', f'{stats["departments"]}个部门'],
            ['岗位类型', f'{stats["positions"]}种职位'],
            [],
            ['岗位分类分布：'],
            ['A类（核心岗位）', f'{stats["a_class"]}人', f'{stats["a_class"]/stats["total"]*100:.1f}%'],
            ['M类（中坚力量）', f'{stats["m_class"]}人', f'{stats["m_class"]/stats["total"]*100:.1f}%'],
            ['E类（专业效能）', f'{stats["e_class"]}人', f'{stats["e_class"]/stats["total"]*100:.1f}%'],
            [],
            ['核心盘点结论：'],
            ['1', f'人岗匹配度：完全匹配{stats["full_match"]}人({stats["full_match_rate"]:.1f}%)，不匹配{stats["no_match"]}人({stats["no_match_rate"]:.1f}%)'],
            ['2', f'人才分层：高潜人才{stats["high_potential"]}人，核心骨干{stats["core_backbone"]}人，待优化{stats["optimization"]}人'],
            ['3', f'风险预警：离职风险{stats["risk"]}人，学历门槛未通过{stats["edu_fail"]}人'],
        ])
        overview.to_excel(writer, sheet_name='一、盘点概览', index=False, header=False)
        
        dept_stats = df.groupby(['部门', '职位']).agg({
            '工号': 'count', '职级': lambda x: ', '.join(sorted(x.unique())), '岗位分类': 'first'
        }).reset_index()
        dept_stats.columns = ['部门', '职位', '人数', '职级分布', '岗位分类']
        dept_stats.to_excel(writer, sheet_name='二、部门职位分布', index=False)
        
        level_order = ['L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'M4', '培训生', '外包']
        level_data = []
        for level in level_order:
            level_df = df[df['职级'] == level]
            if len(level_df) > 0:
                level_data.append([level, len(level_df), f'{len(level_df)/stats["total"]*100:.1f}%', ', '.join(level_df['部门'].unique()[:5])])
        pd.DataFrame(level_data, columns=['职级', '人数', '占比', '主要分布部门']).to_excel(writer, sheet_name='三、职级体系分析', index=False)
        
        edu_fail_df = df[~df['学历门槛通过']][['姓名', '部门', '岗位', '岗位分类', '学历-学校类别', '学历档位']]
        edu_data = [['岗位最低学历门槛校验'], [], ['不满足最低学历要求人员清单：'], ['姓名', '部门', '岗位', '岗位分类', '学历-学校类别', '学历档位']]
        for _, r in edu_fail_df.iterrows():
            edu_data.append([r['姓名'], r['部门'], r['岗位'], r['岗位分类'], r['学历-学校类别'], r['学历档位']])
        if len(edu_fail_df) == 0:
            edu_data.append(['✅ 全部通过学历门槛'])
        pd.DataFrame(edu_data).to_excel(writer, sheet_name='五、学历门槛校验', index=False, header=False)
        
        match_data = [['人岗匹配度校验明细：'], [], ['完全匹配', f'{stats["full_match"]}人', f'{stats["full_match_rate"]:.1f}%'], 
                      ['基本匹配', f'{stats["basic_match"]}人', f'{stats["basic_match_rate"]:.1f}%'], 
                      ['不匹配', f'{stats["no_match"]}人', f'{stats["no_match_rate"]:.1f}%'], [], ['不匹配人员清单：'],
                      ['姓名', '部门', '岗位', '岗位分类', '学历档位', '绩效档位', '匹配度']]
        for _, r in df[df['匹配度'] == '不匹配'].iterrows():
            match_data.append([r['姓名'], r['部门'], r['岗位'], r['岗位分类'], r['学历档位'], r['绩效档位'], r['匹配度']])
        pd.DataFrame(match_data).to_excel(writer, sheet_name='六、人岗匹配度', index=False, header=False)
        
        match_summary = []
        for pc in ['A类', 'M类', 'E类']:
            class_df = df[df['岗位分类'] == pc]
            if len(class_df) > 0:
                for mt in ['完全匹配', '基本匹配', '不匹配']:
                    count = len(class_df[class_df['匹配度'] == mt])
                    match_summary.append([pc, mt, count, f'{count/len(class_df)*100:.1f}%'])
        pd.DataFrame(match_summary, columns=['岗位分类', '匹配度', '人数', '占比']).to_excel(writer, sheet_name='七、匹配度汇总', index=False)
        
        hp_df = df[df['高潜人才']][['姓名', '部门', '岗位', '职级', '司龄', '岗位分类']]
        hp_data = [['高潜人才盘点'], [], [f'高潜人才明细（共{len(hp_df)}人）：'], [], ['姓名', '部门', '岗位', '职级', '司龄', '岗位分类']]
        for _, r in hp_df.iterrows():
            hp_data.append([r['姓名'], r['部门'], r['岗位'], r['职级'], f'{r["司龄"]:.1f}年', r['岗位分类']])
        if len(hp_df) == 0:
            hp_data.append(['暂无符合条件的高潜人才'])
        pd.DataFrame(hp_data).to_excel(writer, sheet_name='八、高潜人才', index=False, header=False)
        
        stable_df = df[df['稳定人员']][['姓名', '部门', '岗位', '职级', '司龄', '岗位分类']]
        stable_data = [['稳定人员盘点'], [], [f'稳定人员明细（共{len(stable_df)}人）：'], [], ['姓名', '部门', '岗位', '职级', '司龄', '岗位分类']]
        for _, r in stable_df.iterrows():
            stable_data.append([r['姓名'], r['部门'], r['岗位'], r['职级'], f'{r["司龄"]:.1f}年', r['岗位分类']])
        pd.DataFrame(stable_data).to_excel(writer, sheet_name='九、稳定人员', index=False, header=False)
        
        cb_df = df[df['核心骨干']][['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度']]
        cb_data = [['核心骨干人员盘点'], [], [f'核心骨干人员明细（共{len(cb_df)}人）：'], [], ['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度']]
        for _, r in cb_df.iterrows():
            cb_data.append([r['姓名'], r['部门'], r['岗位'], r['职级'], r['岗位分类'], r['匹配度']])
        pd.DataFrame(cb_data).to_excel(writer, sheet_name='十、核心骨干', index=False, header=False)
        
        at_df = df[df['待关注']][['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度']]
        at_data = [['待关注人员盘点'], [], [f'待关注人员明细（共{len(at_df)}人）：'], [], ['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度', '风险等级']]
        for _, r in at_df.iterrows():
            at_data.append([r['姓名'], r['部门'], r['岗位'], r['职级'], r['岗位分类'], r['匹配度'], '中'])
        pd.DataFrame(at_data).to_excel(writer, sheet_name='十一、待关注人员', index=False, header=False)
        
        op_df = df[df['待优化']][['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度']]
        op_data = [['待优化人员盘点'], [], [f'待优化人员明细（共{len(op_df)}人）：'], [], ['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度', '优化优先级']]
        for _, r in op_df.iterrows():
            op_data.append([r['姓名'], r['部门'], r['岗位'], r['职级'], r['岗位分类'], r['匹配度'], '高'])
        pd.DataFrame(op_data).to_excel(writer, sheet_name='十二、待优化人员', index=False, header=False)
        
        risk_df = df[df['离职风险']][['姓名', '部门', '岗位', '职级', '岗位分类', '司龄']]
        risk_data = [['离职高风险预警'], [], [f'离职高风险人员清单（共{len(risk_df)}人）：'], [], ['姓名', '部门', '岗位', '职级', '岗位分类', '司龄', '风险等级', '干预建议']]
        for _, r in risk_df.iterrows():
            risk_level = '高' if r['岗位分类'] == 'A类' else '中'
            risk_data.append([r['姓名'], r['部门'], r['岗位'], r['职级'], r['岗位分类'], f'{r["司龄"]:.1f}年', risk_level, '定期沟通，了解诉求'])
        pd.DataFrame(risk_data).to_excel(writer, sheet_name='十三、离职高风险', index=False, header=False)
        
        suggestions = []
        if stats['edu_fail'] > 0:
            suggestions.append(f'1. 学历门槛优化：{stats["edu_fail"]}人未达到岗位最低学历要求，建议制定针对性培养计划或调整岗位定位')
        if stats['high_potential'] > 0:
            suggestions.append(f'2. 高潜人才培养：识别出{stats["high_potential"]}名高潜人才，建议建立专项培养机制，配备导师，加速成长')
        if stats['risk'] > 0:
            suggestions.append(f'3. 人才保留策略：{stats["risk"]}人存在离职风险，建议开展一对一沟通，了解诉求，制定保留方案')
        if stats['optimization'] > 0:
            suggestions.append(f'4. 绩效改进计划：{stats["optimization"]}人列入待优化名单，建议制定PIP，明确改进目标和时间节点')
        suggestions.append('5. 梯队建设：建议通过内部培养或外部引进方式，完善各岗位梯队结构')
        suggestion_data = [['核心优化建议'], [], ['基于本次盘点结果，提出以下优化建议：'], []] + [[s] for s in suggestions]
        pd.DataFrame(suggestion_data).to_excel(writer, sheet_name='十五、优化建议', index=False, header=False)
        
        export_cols = ['工号', '姓名', '部门', '岗位', '职位', '职级', '岗位分类', '学历档位', '绩效档位', '匹配度', '司龄', '高潜人才', '稳定人员', '核心骨干', '待关注', '待优化', '离职风险']
        available_cols = [c for c in export_cols if c in df.columns]
        df[available_cols].to_excel(writer, sheet_name='十六、完整盘点数据', index=False)
        
    output.seek(0)
    return output

def generate_pdf_report(df, stats):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        import os
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        font_paths = [
            os.path.join(script_dir, 'NotoSansSC-Regular.otf'),
            'NotoSansSC-Regular.otf',
            '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
            '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
            '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
            '/System/Library/Fonts/STHeiti Medium.ttc',
            '/System/Library/Fonts/Hiragino Sans GB.ttc',
        ]
        chinese_font = 'Helvetica'
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    pdfmetrics.registerFont(TTFont('ChineseFont', fp))
                    chinese_font = 'ChineseFont'
                    break
                except Exception as e:
                    continue
        
        output = BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName=chinese_font, fontSize=22, alignment=1, spaceAfter=20, textColor=colors.HexColor('#2C3E50'))
        heading_style = ParagraphStyle('Heading', parent=styles['Heading2'], fontName=chinese_font, fontSize=14, spaceAfter=10, spaceBefore=15, textColor=colors.HexColor('#2C3E50'))
        subheading_style = ParagraphStyle('SubHeading', parent=styles['Normal'], fontName=chinese_font, fontSize=12, spaceAfter=8, spaceBefore=10, textColor=colors.HexColor('#34495E'))
        normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontName=chinese_font, fontSize=10, leading=14)
        highlight_style = ParagraphStyle('Highlight', parent=styles['Normal'], fontName=chinese_font, fontSize=11, textColor=colors.HexColor('#E74C3C'), leading=15)
        
        story = []
        story.append(Paragraph('产研团队全量人才盘点报告', title_style))
        story.append(Paragraph(f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}', normal_style))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('一、盘点概览', heading_style))
        overview_data = [
            ['盘点指标', '数据详情'],
            ['盘点总人数', f'{stats["total"]}人（正式员工{stats["formal"]}人）'],
            ['部门覆盖', f'{stats["departments"]}个部门'],
            ['岗位类型', f'{stats["positions"]}种职位'],
            ['盘点范围', '产研团队全体员工'],
        ]
        t = Table(overview_data, colWidths=[4*cm, 12*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498DB')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('BACKGROUND', (0, 1), (0, -1), colors.HexColor('#EBF5FB')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), chinese_font),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ]))
        story.append(t)
        story.append(Spacer(1, 15))
        
        story.append(Paragraph('岗位分类分布', subheading_style))
        class_data = [
            ['岗位分类', '定义说明', '人数', '占比'],
            ['A类（核心岗位）', '对学历和绩效要求最高', f'{stats["a_class"]}人', f'{stats["a_class"]/stats["total"]*100:.1f}%'],
            ['M类（中坚力量）', '承上启下的关键角色', f'{stats["m_class"]}人', f'{stats["m_class"]/stats["total"]*100:.1f}%'],
            ['E类（专业效能）', '专业执行类岗位', f'{stats["e_class"]}人', f'{stats["e_class"]/stats["total"]*100:.1f}%'],
        ]
        t2 = Table(class_data, colWidths=[3.5*cm, 6*cm, 3*cm, 3.5*cm])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C3E50')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 1), (1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), chinese_font),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ]))
        story.append(t2)
        story.append(Spacer(1, 15))
        
        story.append(Paragraph('人岗匹配度概览', subheading_style))
        match_data = [
            ['匹配度', '人数', '占比', '说明'],
            ['完全匹配', f'{stats["full_match"]}人', f'{stats["full_match_rate"]:.1f}%', '学历+绩效双优'],
            ['基本匹配', f'{stats["basic_match"]}人', f'{stats["basic_match_rate"]:.1f}%', '满足基本要求'],
            ['不匹配', f'{stats["no_match"]}人', f'{stats["no_match_rate"]:.1f}%', '需重点关注'],
        ]
        t3 = Table(match_data, colWidths=[3*cm, 3*cm, 3*cm, 7*cm])
        t3.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#27AE60')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor('#FADBD8')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (3, 1), (3, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), chinese_font),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ]))
        story.append(t3)
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('二、人才盘点结果', heading_style))
        talent_data = [
            ['人才分类', '人数', '占比', '特征说明'],
            ['高潜人才', f'{stats["high_potential"]}人', f'{stats["high_potential_rate"]:.1f}%', '司龄≤1年+近2期绩效全A/S'],
            ['稳定人员', f'{stats["stable"]}人', f'{stats["stable_rate"]:.1f}%', '司龄≥0.5年+4期绩效全B+以上'],
            ['核心骨干', f'{stats["core_backbone"]}人', f'{stats["core_backbone_rate"]:.1f}%', 'A/M类+绩效优+L6以上+完全匹配'],
            ['待关注人员', f'{stats["attention"]}人', f'{stats["attention_rate"]:.1f}%', '绩效波动或基本匹配'],
            ['待优化人员', f'{stats["optimization"]}人', f'{stats["optimization_rate"]:.1f}%', '绩效C或人岗不匹配'],
            ['离职高风险', f'{stats["risk"]}人', f'{stats["risk_rate"]:.1f}%', '绩效下滑或A类待优化'],
        ]
        t4 = Table(talent_data, colWidths=[3*cm, 2.5*cm, 2.5*cm, 8*cm])
        t4.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#8E44AD')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#D5F5E3')),
            ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#D6EAF8')),
            ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor('#FCF3CF')),
            ('BACKGROUND', (0, 4), (-1, 4), colors.HexColor('#FAE5D3')),
            ('BACKGROUND', (0, 5), (-1, 5), colors.HexColor('#FADBD8')),
            ('BACKGROUND', (0, 6), (-1, 6), colors.HexColor('#F5B7B1')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (3, 1), (3, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), chinese_font),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ]))
        story.append(t4)
        story.append(Spacer(1, 15))
        
        if stats['edu_fail'] > 0:
            story.append(Paragraph(f'⚠️ 学历门槛预警：{stats["edu_fail"]}人未达到岗位最低学历要求', highlight_style))
            story.append(Spacer(1, 10))
        
        story.append(PageBreak())
        
        story.append(Paragraph('三、核心优化建议', heading_style))
        suggestions = []
        if stats['edu_fail'] > 0:
            suggestions.append(f'【学历门槛优化】{stats["edu_fail"]}人未达到岗位最低学历要求，建议制定针对性培养计划或调整岗位定位。')
        if stats['high_potential'] > 0:
            suggestions.append(f'【高潜人才培养】识别出{stats["high_potential"]}名高潜人才，建议建立专项培养机制，配备导师，加速成长。')
        if stats['core_backbone'] > 0:
            suggestions.append(f'【核心骨干保留】{stats["core_backbone"]}名核心骨干是团队关键资产，建议制定激励方案确保留存。')
        if stats['risk'] > 0:
            suggestions.append(f'【人才保留策略】{stats["risk"]}人存在离职风险，建议开展一对一沟通，了解诉求，制定保留方案。')
        if stats['attention'] > 0:
            suggestions.append(f'【绩效关注】{stats["attention"]}人列入待关注名单，建议定期跟进绩效表现，及时干预。')
        if stats['optimization'] > 0:
            suggestions.append(f'【绩效改进计划】{stats["optimization"]}人列入待优化名单，建议制定PIP，明确改进目标和时间节点。')
        suggestions.append('【梯队建设】建议通过内部培养或外部引进方式，完善各岗位梯队结构，确保关键岗位有人可用。')
        suggestions.append('【定期盘点】建议每半年进行一次人才盘点，及时掌握团队动态，优化人才配置。')
        for i, s in enumerate(suggestions[:8], 1):
            story.append(Paragraph(f'{i}. {s}', normal_style))
            story.append(Spacer(1, 8))
        
        story.append(Spacer(1, 20))
        story.append(Paragraph('四、重点人员清单', heading_style))
        
        if stats['optimization'] > 0:
            story.append(Paragraph('待优化人员清单（需重点关注）', subheading_style))
            opt_df = df[df['待优化']][['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度']].head(10)
            opt_data = [['姓名', '部门', '岗位', '职级', '岗位分类', '匹配度']]
            for _, r in opt_df.iterrows():
                opt_data.append([str(r['姓名']), str(r['部门'])[:8], str(r['岗位'])[:10], str(r['职级']), str(r['岗位分类']), str(r['匹配度'])])
            t_opt = Table(opt_data, colWidths=[2.5*cm, 3*cm, 3.5*cm, 2*cm, 2.5*cm, 2.5*cm])
            t_opt.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E74C3C')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), chinese_font),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
            ]))
            story.append(t_opt)
            story.append(Spacer(1, 15))
        
        if stats['risk'] > 0:
            story.append(Paragraph('离职高风险人员清单', subheading_style))
            risk_df = df[df['离职风险']][['姓名', '部门', '岗位', '职级', '岗位分类', '司龄']].head(10)
            risk_data = [['姓名', '部门', '岗位', '职级', '岗位分类', '司龄']]
            for _, r in risk_df.iterrows():
                risk_data.append([str(r['姓名']), str(r['部门'])[:8], str(r['岗位'])[:10], str(r['职级']), str(r['岗位分类']), f'{r["司龄"]:.1f}年'])
            t_risk = Table(risk_data, colWidths=[2.5*cm, 3*cm, 3.5*cm, 2*cm, 2.5*cm, 2.5*cm])
            t_risk.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#C0392B')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), chinese_font),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
            ]))
            story.append(t_risk)
            story.append(Spacer(1, 15))
        
        if stats['high_potential'] > 0:
            story.append(Paragraph('高潜人才清单', subheading_style))
            hp_df = df[df['高潜人才']][['姓名', '部门', '岗位', '职级', '岗位分类', '司龄']].head(10)
            hp_data = [['姓名', '部门', '岗位', '职级', '岗位分类', '司龄']]
            for _, r in hp_df.iterrows():
                hp_data.append([str(r['姓名']), str(r['部门'])[:8], str(r['岗位'])[:10], str(r['职级']), str(r['岗位分类']), f'{r["司龄"]:.1f}年'])
            t_hp = Table(hp_data, colWidths=[2.5*cm, 3*cm, 3.5*cm, 2*cm, 2.5*cm, 2.5*cm])
            t_hp.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#27AE60')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), chinese_font),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
            ]))
            story.append(t_hp)
        
        story.append(Spacer(1, 30))
        footer_style = ParagraphStyle('Footer', parent=styles['Normal'], fontName=chinese_font, fontSize=9, textColor=colors.HexColor('#95A5A6'), alignment=1)
        story.append(Paragraph('—— 报告结束 ——', footer_style))
        story.append(Paragraph(f'生成时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', footer_style))
        
        doc.build(story)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"PDF生成失败: {str(e)}")
        return None

with st.sidebar:
    st.title("📊 产研团队人才盘点系统")
    st.markdown("---")
    page = st.radio("功能导航", ["🏠 首页", "📤 数据上传", "📊 盘点分析", "📄 报告导出"], label_visibility="collapsed")
    st.markdown("---")
    st.markdown(f"<div style='text-align:center;color:#666;font-size:0.8rem;'>版本：v2.0 Web版<br>{datetime.now().strftime('%Y-%m-%d')}</div>", unsafe_allow_html=True)

if page == "🏠 首页":
    st.markdown("<h1 style='text-align:center;color:#1E3A8A;'>产研团队全量人才盘点系统</h1>", unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("👥 总人数", value="--")
    with col2: st.metric("🏢 部门数", value="--")
    with col3: st.metric("✅ 匹配率", value="--")
    with col4: st.metric("⚠️ 风险人数", value="--")
    st.markdown("---")
    st.markdown("""
    ### 🎯 系统功能
    | 功能模块 | 说明 |
    |---------|------|
    | 📤 数据上传 | 上传员工基础信息表，自动解析数据 |
    | 📊 盘点分析 | 执行人才盘点，生成详细分析结果 |
    | 📄 报告导出 | 导出PDF/Excel格式报告（与本地版一致） |
    
    ### 📝 使用步骤
    1. 点击左侧 **"数据上传"** 上传员工信息表
    2. 点击 **"盘点分析"** 执行盘点
    3. 点击 **"报告导出"** 下载PDF/Excel报告
    
    ### 📋 输出报告说明
    **Excel报告（16个工作表）：**
    - 一、盘点概览
    - 二、部门职位分布
    - 三、职级体系分析
    - 五、学历门槛校验
    - 六、人岗匹配度
    - 七、匹配度汇总
    - 八、高潜人才
    - 九、稳定人员
    - 十、核心骨干
    - 十一、待关注人员
    - 十二、待优化人员
    - 十三、离职高风险
    - 十五、优化建议
    - 十六、完整盘点数据
    
    **PDF报告（汇报版）：**
    - 盘点概览（总人数、部门覆盖、岗位分类分布）
    - 人才盘点结果（高潜人才、稳定人员、核心骨干等）
    - 核心优化建议
    - 重点人员清单
    """)

elif page == "📤 数据上传":
    st.header("📤 数据上传")
    st.markdown("请上传Excel格式的员工基础信息表")
    uploaded_file = st.file_uploader("选择文件", type=['xlsx', 'xls'])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ 文件上传成功！共读取 {len(df)} 条记录")
            st.dataframe(df.head(10), use_container_width=True)
            col1, col2, col3 = st.columns(3)
            with col1: st.metric("总人数", f"{len(df)}人")
            with col2: st.metric("部门数", f"{df['部门'].nunique()}个")
            with col3: st.metric("岗位数", f"{df['职位'].nunique()}种")
            if st.button("✅ 确认使用此数据", type="primary"):
                st.session_state.df = df
                st.session_state.result_df = None
                st.session_state.stats = None
                st.success("✅ 数据已加载，请前往'盘点分析'执行盘点")
        except Exception as e:
            st.error(f"❌ 文件读取失败：{str(e)}")

elif page == "📊 盘点分析":
    st.header("📊 盘点分析")
    if st.session_state.df is None:
        st.warning("⚠️ 请先上传员工数据")
    else:
        st.markdown(f"**当前数据：** {len(st.session_state.df)} 条记录")
        if st.button("🚀 执行盘点分析", type="primary"):
            with st.spinner("正在执行盘点分析..."):
                result_df = perform_review(st.session_state.df)
                st.session_state.result_df = result_df
                st.session_state.stats = calculate_stats(result_df)
            st.success("✅ 盘点分析完成！")
            st.rerun()
        
        if st.session_state.result_df is not None:
            stats = st.session_state.stats
            df = st.session_state.result_df
            st.markdown("### 📈 盘点结果概览")
            col1, col2, col3, col4 = st.columns(4)
            with col1: st.metric("总人数", f"{stats['total']}人")
            with col2: st.metric("完全匹配", f"{stats['full_match']}人", f"{stats['full_match_rate']:.1f}%")
            with col3: st.metric("待优化", f"{stats['optimization']}人", f"{stats['optimization_rate']:.1f}%")
            with col4: st.metric("离职风险", f"{stats['risk']}人", f"{stats['risk_rate']:.1f}%")
            
            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### 人才分层分布")
                layer_counts = df['人才分层'].value_counts()
                fig = px.pie(values=layer_counts.values, names=layer_counts.index, title='人才分层分布', hole=0.4)
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                st.markdown("#### 人岗匹配度分布")
                match_counts = df['匹配度'].value_counts()
                fig2 = px.pie(values=match_counts.values, names=match_counts.index, title='人岗匹配度分布', 
                             color=match_counts.index, color_discrete_map={'完全匹配': '#2ECC71', '基本匹配': '#F39C12', '不匹配': '#E74C3C'})
                st.plotly_chart(fig2, use_container_width=True)
            
            st.markdown("---")
            st.markdown("#### 人才分类统计")
            talent_stats = pd.DataFrame({
                '分类': ['高潜人才', '稳定人员', '核心骨干', '待关注', '待优化', '离职风险'],
                '人数': [stats['high_potential'], stats['stable'], stats['core_backbone'], stats['attention'], stats['optimization'], stats['risk']],
                '占比': [f"{stats['high_potential_rate']:.1f}%", f"{stats['stable_rate']:.1f}%", f"{stats['core_backbone_rate']:.1f}%", f"{stats['attention_rate']:.1f}%", f"{stats['optimization_rate']:.1f}%", f"{stats['risk_rate']:.1f}%"]
            })
            st.dataframe(talent_stats, use_container_width=True, hide_index=True)

elif page == "📄 报告导出":
    st.header("📄 报告导出")
    if st.session_state.result_df is None:
        st.warning("⚠️ 请先完成盘点分析")
    else:
        df = st.session_state.result_df
        stats = st.session_state.stats
        st.markdown("### 选择导出格式")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📊 Excel报告（完整版）")
            st.markdown("包含16个工作表的详细盘点报告，与本地版一致")
            if st.button("导出Excel", type="primary", key="excel_btn"):
                with st.spinner("正在生成Excel报告..."):
                    excel_output = generate_excel_report(df, stats)
                st.download_button(
                    "📥 下载 Excel 报告", 
                    data=excel_output, 
                    file_name=f"产研团队人才盘点完整报告_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col2:
            st.markdown("#### 📄 PDF报告（汇报版）")
            st.markdown("适合汇报场景的PDF格式报告，与本地版一致")
            if st.button("导出PDF", type="primary", key="pdf_btn"):
                with st.spinner("正在生成PDF报告..."):
                    pdf_output = generate_pdf_report(df, stats)
                    if pdf_output:
                        st.download_button(
                            "📥 下载 PDF 报告",
                            data=pdf_output,
                            file_name=f"人才盘点报告-汇报版_{datetime.now().strftime('%Y%m%d')}.pdf",
                            mime="application/pdf"
                        )
                    else:
                        st.error("PDF生成失败，请检查系统字体配置")
        
        st.markdown("---")
        st.markdown("#### 📊 数据预览")
        st.dataframe(df.head(20), use_container_width=True)
