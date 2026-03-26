import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
import plotly.graph_objects as go
import plotly.express as px
import os
import urllib.request

st.set_page_config(
    page_title="产研团队人才盘点系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

FONT_FILE = "NotoSansSC-Regular.otf"
FONT_URL = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/SimplifiedChinese/NotoSansSC-Regular.otf"

def get_font_path():
    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
    font_path = os.path.join(script_dir, FONT_FILE)
    if not os.path.exists(font_path):
        font_path = FONT_FILE
    if not os.path.exists(font_path):
        try:
            urllib.request.urlretrieve(FONT_URL, FONT_FILE)
            font_path = FONT_FILE
        except:
            pass
    return font_path if os.path.exists(font_path) else None

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
        
        export_cols = ['工号', '姓名', '部门', '岗位', '职位', '职级', '岗位分类', '学历档位', '绩效档位', '匹配度', '司龄', '高潜人才', '稳定人员', '核心骨干', '待关注', '待优化', '离职风险']
        available_cols = [c for c in export_cols if c in df.columns]
        df[available_cols].to_excel(writer, sheet_name='完整盘点数据', index=False)
        
    output.seek(0)
    return output

def generate_pdf_report(df, stats):
    try:
        from fpdf import FPDF
        
        font_path = get_font_path()
        output = BytesIO()
        
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        if font_path:
            pdf.add_font('NotoSC', '', font_path)
            pdf.set_font('NotoSC', '', 18)
            pdf.cell(0, 12, txt='产研团队人才盘点报告', ln=True, align='C')
            pdf.ln(3)
            pdf.set_font('NotoSC', '', 10)
            pdf.cell(0, 8, txt=f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}', ln=True, align='C')
        else:
            pdf.set_font('Helvetica', '', 18)
            pdf.cell(0, 12, txt='Talent Review Report', ln=True, align='C')
            pdf.ln(3)
            pdf.set_font('Helvetica', '', 10)
            pdf.cell(0, 8, txt=f'Date: {datetime.now().strftime("%Y-%m-%d")}', ln=True, align='C')
        
        pdf.ln(10)
        
        font_name = 'NotoSC' if font_path else 'Helvetica'
        
        pdf.set_font(font_name, '', 14)
        pdf.cell(0, 10, txt='一、盘点概览' if font_path else '1. Overview', ln=True)
        pdf.ln(3)
        
        pdf.set_font(font_name, '', 10)
        overview_items = [
            (f'盘点总人数：{stats["total"]}人' if font_path else f'Total: {stats["total"]}'),
            (f'部门覆盖：{stats["departments"]}个' if font_path else f'Departments: {stats["departments"]}'),
            (f'完全匹配：{stats["full_match"]}人 ({stats["full_match_rate"]:.1f}%)' if font_path else f'Full Match: {stats["full_match"]} ({stats["full_match_rate"]:.1f}%)'),
            (f'待优化：{stats["optimization"]}人' if font_path else f'Need Optimization: {stats["optimization"]}'),
            (f'离职风险：{stats["risk"]}人' if font_path else f'Resignation Risk: {stats["risk"]}'),
        ]
        for item in overview_items:
            pdf.cell(0, 7, txt=item, ln=True)
        
        pdf.ln(8)
        pdf.set_font(font_name, '', 14)
        pdf.cell(0, 10, txt='二、人才分层统计' if font_path else '2. Talent Categories', ln=True)
        pdf.ln(3)
        
        pdf.set_font(font_name, '', 10)
        talent_items = [
            (f'高潜人才：{stats["high_potential"]}人 ({stats["high_potential_rate"]:.1f}%)' if font_path else f'High Potential: {stats["high_potential"]} ({stats["high_potential_rate"]:.1f}%)'),
            (f'稳定人员：{stats["stable"]}人 ({stats["stable_rate"]:.1f}%)' if font_path else f'Stable: {stats["stable"]} ({stats["stable_rate"]:.1f}%)'),
            (f'核心骨干：{stats["core_backbone"]}人 ({stats["core_backbone_rate"]:.1f}%)' if font_path else f'Core Backbone: {stats["core_backbone"]} ({stats["core_backbone_rate"]:.1f}%)'),
            (f'待关注：{stats["attention"]}人 ({stats["attention_rate"]:.1f}%)' if font_path else f'Need Attention: {stats["attention"]} ({stats["attention_rate"]:.1f}%)'),
            (f'待优化：{stats["optimization"]}人 ({stats["optimization_rate"]:.1f}%)' if font_path else f'Need Optimization: {stats["optimization"]} ({stats["optimization_rate"]:.1f}%)'),
            (f'离职风险：{stats["risk"]}人 ({stats["risk_rate"]:.1f}%)' if font_path else f'Resignation Risk: {stats["risk"]} ({stats["risk_rate"]:.1f}%)'),
        ]
        for item in talent_items:
            pdf.cell(0, 7, txt=item, ln=True)
        
        pdf.ln(8)
        pdf.set_font(font_name, '', 14)
        pdf.cell(0, 10, txt='三、岗位分类分布' if font_path else '3. Position Classes', ln=True)
        pdf.ln(3)
        
        pdf.set_font(font_name, '', 10)
        class_items = [
            (f'A类（核心岗位）：{stats["a_class"]}人' if font_path else f'Class A: {stats["a_class"]}'),
            (f'M类（中坚力量）：{stats["m_class"]}人' if font_path else f'Class M: {stats["m_class"]}'),
            (f'E类（专业效能）：{stats["e_class"]}人' if font_path else f'Class E: {stats["e_class"]}'),
        ]
        for item in class_items:
            pdf.cell(0, 7, txt=item, ln=True)
        
        pdf.add_page()
        pdf.set_font(font_name, '', 14)
        pdf.cell(0, 10, txt='四、核心优化建议' if font_path else '4. Recommendations', ln=True)
        pdf.ln(3)
        
        pdf.set_font(font_name, '', 10)
        recs = [
            (f'1. 高潜人才培养：识别出{stats["high_potential"]}名高潜人才，建议建立专项培养机制。' if font_path else f'1. High potential training: {stats["high_potential"]} talents identified.'),
            (f'2. 核心骨干保留：{stats["core_backbone"]}名核心骨干是团队关键资产。' if font_path else f'2. Core backbone retention: {stats["core_backbone"]} key assets.'),
            (f'3. 待优化人员处理：{stats["optimization"]}人需制定绩效改进计划。' if font_path else f'3. Performance improvement: {stats["optimization"]} people need PIP.'),
            (f'4. 离职风险防控：{stats["risk"]}人存在离职风险，建议开展沟通。' if font_path else f'4. Retention strategy: {stats["risk"]} people at risk.'),
            (f'5. 梯队建设：建议完善各岗位梯队结构。' if font_path else '5. Team structure: Improve echelon structure.'),
        ]
        for rec in recs:
            pdf.multi_cell(0, 7, txt=rec, ln=True)
            pdf.ln(2)
        
        pdf_content = pdf.output(dest='S')
        if isinstance(pdf_content, str):
            output.write(pdf_content.encode('latin-1'))
        else:
            output.write(pdf_content)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"PDF生成失败: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

with st.sidebar:
    st.title("📊 产研团队人才盘点系统")
    st.markdown("---")
    page = st.radio("功能导航", ["🏠 首页", "📤 数据上传", "📊 盘点分析", "📄 报告导出"], label_visibility="collapsed")
    st.markdown("---")
    font_status = "✅ 中文字体已加载" if get_font_path() else "⚠️ 中文字体未找到"
    st.markdown(f"<div style='text-align:center;color:#666;font-size:0.8rem;'>版本：v2.1<br>{font_status}</div>", unsafe_allow_html=True)

if page == "🏠 首页":
    st.markdown("<h1 style='text-align:center;color:#1E3A8A;'>产研团队全量人才盘点系统</h1>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("""
    ### 🎯 系统功能
    - 📤 数据上传：上传员工基础信息表
    - 📊 盘点分析：执行人才盘点
    - 📄 报告导出：导出PDF/Excel报告
    
    ### 📝 使用步骤
    1. 点击左侧 **"数据上传"** 上传员工信息表
    2. 点击 **"盘点分析"** 执行盘点
    3. 点击 **"报告导出"** 下载报告
    """)

elif page == "📤 数据上传":
    st.header("📤 数据上传")
    uploaded_file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'])
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ 成功读取 {len(df)} 条记录")
            st.dataframe(df.head(10), use_container_width=True)
            if st.button("✅ 确认使用", type="primary"):
                st.session_state.df = df
                st.session_state.result_df = None
                st.session_state.stats = None
                st.success("✅ 数据已加载")
        except Exception as e:
            st.error(f"❌ 读取失败：{e}")

elif page == "📊 盘点分析":
    st.header("📊 盘点分析")
    if st.session_state.df is None:
        st.warning("⚠️ 请先上传数据")
    else:
        if st.button("🚀 执行盘点", type="primary"):
            with st.spinner("分析中..."):
                result = perform_review(st.session_state.df)
                st.session_state.result_df = result
                st.session_state.stats = calculate_stats(result)
            st.success("✅ 完成！")
            st.rerun()
        
        if st.session_state.result_df is not None:
            s = st.session_state.stats
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("总人数", s['total'])
            c2.metric("完全匹配", f"{s['full_match']}人", f"{s['full_match_rate']:.1f}%")
            c3.metric("待优化", f"{s['optimization']}人")
            c4.metric("离职风险", f"{s['risk']}人")
            
            st.dataframe(st.session_state.result_df.head(10))

elif page == "📄 报告导出":
    st.header("📄 报告导出")
    if st.session_state.result_df is None:
        st.warning("⚠️ 请先完成盘点")
    else:
        df = st.session_state.result_df
        stats = st.session_state.stats
        
        font_path = get_font_path()
        if font_path:
            st.success(f"✅ 检测到中文字体，PDF将显示中文")
        else:
            st.warning("⚠️ 未检测到中文字体，PDF将显示英文")
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### 📊 Excel报告")
            if st.button("导出Excel", type="primary"):
                excel = generate_excel_report(df, stats)
                st.download_button("下载Excel", excel, f"人才盘点报告_{datetime.now().strftime('%Y%m%d')}.xlsx")
        
        with c2:
            st.markdown("### 📄 PDF报告")
            if st.button("导出PDF", type="primary"):
                pdf = generate_pdf_report(df, stats)
                if pdf:
                    filename = f"人才盘点报告-汇报版_{datetime.now().strftime('%Y%m%d')}.pdf" if font_path else f"Talent_Review_Report_{datetime.now().strftime('%Y%m%d')}.pdf"
                    st.download_button("下载PDF", pdf, filename, "application/pdf")
