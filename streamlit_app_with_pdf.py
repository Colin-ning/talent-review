import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
import base64

st.set_page_config(
    page_title="产研团队人才盘点系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header { font-size: 2rem; color: #1E3A8A; text-align: center; margin-bottom: 1.5rem; }
    .stMetric > div { background-color: #f0f2f6; padding: 10px; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

if 'df' not in st.session_state:
    st.session_state.df = None
if 'result_df' not in st.session_state:
    st.session_state.result_df = None
if 'stats' not in st.session_state:
    st.session_state.stats = None

def get_education_tier(row):
    school_type = str(row.get('学历-学校类别', ''))
    qs_rank = row.get('海外高校QS排名', None)
    if pd.notna(qs_rank) and qs_rank <= 300: return '一档'
    if school_type in ['C9', '985']: return '一档'
    if pd.notna(qs_rank) and qs_rank <= 500: return '二档'
    if school_type in ['211', '双一流', '类211']: return '二档'
    if school_type in ['一本', '二本', '三本', '大专']: return '三档'
    return '三档'

def get_performance_tier(row):
    cols = ['2024上半年绩效结果', '2024年度绩效结果', '2025上半年绩效结果', '2025年度绩效结果']
    perfs = [row[c] for c in cols if c in row.index and pd.notna(row[c]) and row[c] != '-']
    if len(perfs) == 0: return '二档'
    sa_count = sum(1 for p in perfs if p in ['S', 'A'])
    all_bp = all(p in ['S', 'A', 'B+'] for p in perfs)
    if sa_count / len(perfs) > 0.5 and all_bp: return '一档'
    if all_bp: return '二档'
    return '三档'

def check_min_education(row):
    school_type = str(row.get('学历-学校类别', ''))
    qs_rank = row.get('海外高校QS排名', None)
    position_class = str(row.get('岗位分类', ''))
    tier1, tier2 = ['C9', '985'], ['211', '双一流', '类211']
    if position_class == 'A类':
        return school_type in tier1 + tier2 or (pd.notna(qs_rank) and qs_rank <= 500)
    elif position_class == 'M类':
        return school_type in tier1 + tier2 + ['一本'] or (pd.notna(qs_rank) and qs_rank <= 1000)
    elif position_class == 'E类':
        return school_type in tier1 + tier2 + ['一本', '二本']
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

def perform_review(df):
    df = df.copy()
    df['学历档位'] = df.apply(get_education_tier, axis=1)
    df['绩效档位'] = df.apply(get_performance_tier, axis=1)
    df['学历门槛通过'] = df.apply(check_min_education, axis=1)
    df['匹配度'] = df.apply(check_match, axis=1)
    today = datetime.now()
    df['司龄'] = df['入职时间'].apply(lambda x: (today - pd.to_datetime(x)).days / 365 if pd.notna(x) else 0)
    df['高潜人才'] = df.apply(lambda r: r.get('司龄', 0) <= 1 and r.get('职级', '') in ['L3', 'L4', 'L5', 'L6', '培训生'], axis=1)
    df['稳定人员'] = df.apply(lambda r: r.get('司龄', 0) >= 0.5, axis=1)
    df['核心骨干'] = df.apply(lambda r: r.get('岗位分类', '') in ['A类', 'M类'] and r.get('匹配度', '') == '完全匹配' and r.get('职级', '') in ['L6', 'L7', 'L8', 'L9', 'M4'], axis=1)
    df['待关注'] = df.apply(lambda r: r.get('匹配度', '') == '基本匹配', axis=1)
    df['待优化'] = df['匹配度'] == '不匹配'
    df['离职风险'] = df.apply(lambda r: 2 <= r.get('司龄', 0) <= 3, axis=1)
    def get_layer(r):
        if r['核心骨干']: return '核心骨干'
        if r['高潜人才']: return '高潜人才'
        if r['稳定人员']: return '稳定人员'
        if r['待优化']: return '待优化'
        if r['待关注']: return '待关注'
        return '其他'
    df['人才分层'] = df.apply(get_layer, axis=1)
    return df

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
        
        font_paths = [
            '/System/Library/Fonts/STHeiti Medium.ttc',
            '/System/Library/Fonts/STHeiti Light.ttc',
            '/System/Library/Fonts/Hiragino Sans GB.ttc',
            '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        ]
        chinese_font = None
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    pdfmetrics.registerFont(TTFont('ChineseFont', fp, subfontIndex=0))
                    chinese_font = 'ChineseFont'
                    break
                except:
                    continue
        
        if not chinese_font:
            return None
        
        output = BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName=chinese_font, fontSize=20, alignment=1, spaceAfter=20)
        heading_style = ParagraphStyle('Heading', parent=styles['Heading2'], fontName=chinese_font, fontSize=14, spaceAfter=10, spaceBefore=15)
        normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontName=chinese_font, fontSize=10, leading=14)
        
        story = []
        story.append(Paragraph('产研团队人才盘点报告', title_style))
        story.append(Paragraph(f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}', normal_style))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('一、盘点概览', heading_style))
        overview_data = [
            ['指标', '数据'],
            ['盘点总人数', f'{stats["total"]}人'],
            ['部门覆盖', f'{stats["departments"]}个'],
            ['完全匹配', f'{stats["full_match"]}人 ({stats["full_match_rate"]:.1f}%)'],
            ['基本匹配', f'{stats["basic_match"]}人 ({stats["basic_match_rate"]:.1f}%)'],
            ['不匹配', f'{stats["no_match"]}人 ({stats["no_match_rate"]:.1f}%)'],
        ]
        t = Table(overview_data, colWidths=[6*cm, 10*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498DB')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), chinese_font),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(t)
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('二、人才分层统计', heading_style))
        talent_data = [
            ['人才分类', '人数', '占比'],
            ['高潜人才', f'{stats["high_potential"]}人', f'{stats["high_potential_rate"]:.1f}%'],
            ['稳定人员', f'{stats["stable"]}人', f'{stats["stable_rate"]:.1f}%'],
            ['核心骨干', f'{stats["core_backbone"]}人', f'{stats["core_backbone_rate"]:.1f}%'],
            ['待关注', f'{stats["attention"]}人', f'{stats["attention_rate"]:.1f}%'],
            ['待优化', f'{stats["optimization"]}人', f'{stats["optimization_rate"]:.1f}%'],
            ['离职风险', f'{stats["risk"]}人', f'{stats["risk_rate"]:.1f}%'],
        ]
        t2 = Table(talent_data, colWidths=[5*cm, 4*cm, 4*cm])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#8E44AD')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), chinese_font),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(t2)
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('三、岗位分类分布', heading_style))
        class_data = [
            ['岗位分类', '人数', '占比'],
            ['A类（核心岗位）', f'{stats["a_class"]}人', f'{stats["a_class"]/stats["total"]*100:.1f}%'],
            ['M类（中坚力量）', f'{stats["m_class"]}人', f'{stats["m_class"]/stats["total"]*100:.1f}%'],
            ['E类（专业效能）', f'{stats["e_class"]}人', f'{stats["e_class"]/stats["total"]*100:.1f}%'],
        ]
        t3 = Table(class_data, colWidths=[5*cm, 4*cm, 4*cm])
        t3.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C3E50')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), chinese_font),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(t3)
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('四、核心优化建议', heading_style))
        suggestions = [
            f'1. 高潜人才培养：识别出{stats["high_potential"]}名高潜人才，建议建立专项培养机制。',
            f'2. 核心骨干保留：{stats["core_backbone"]}名核心骨干是团队关键资产，建议制定激励方案。',
            f'3. 待优化人员处理：{stats["optimization"]}人列入待优化名单，建议制定绩效改进计划。',
            f'4. 离职风险防控：{stats["risk"]}人存在离职风险，建议开展一对一沟通。',
            f'5. 梯队建设：建议通过内部培养或外部引进方式，完善梯队结构。',
        ]
        for s in suggestions:
            story.append(Paragraph(s, normal_style))
            story.append(Spacer(1, 8))
        
        story.append(PageBreak())
        
        story.append(Paragraph('五、重点人员清单', heading_style))
        
        if stats["optimization"] > 0:
            story.append(Paragraph('待优化人员清单', heading_style))
            opt_df = df[df['待优化']][['姓名', '部门', '岗位', '职级', '匹配度']].head(10)
            opt_data = [['姓名', '部门', '岗位', '职级', '匹配度']]
            for _, r in opt_df.iterrows():
                opt_data.append([str(r['姓名']), str(r['部门'])[:10], str(r['岗位'])[:10], str(r['职级']), str(r['匹配度'])])
            t_opt = Table(opt_data, colWidths=[2.5*cm, 4*cm, 4*cm, 2*cm, 2.5*cm])
            t_opt.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E74C3C')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), chinese_font),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            story.append(t_opt)
            story.append(Spacer(1, 15))
        
        if stats["risk"] > 0:
            story.append(Paragraph('离职风险人员清单', heading_style))
            risk_df = df[df['离职风险']][['姓名', '部门', '岗位', '职级', '司龄']].head(10)
            risk_data = [['姓名', '部门', '岗位', '职级', '司龄']]
            for _, r in risk_df.iterrows():
                risk_data.append([str(r['姓名']), str(r['部门'])[:10], str(r['岗位'])[:10], str(r['职级']), f'{r["司龄"]:.1f}年'])
            t_risk = Table(risk_data, colWidths=[2.5*cm, 4*cm, 4*cm, 2*cm, 2.5*cm])
            t_risk.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#C0392B')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), chinese_font),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            story.append(t_risk)
            story.append(Spacer(1, 15))
        
        if stats["high_potential"] > 0:
            story.append(Paragraph('高潜人才清单', heading_style))
            hp_df = df[df['高潜人才']][['姓名', '部门', '岗位', '职级', '司龄']].head(10)
            hp_data = [['姓名', '部门', '岗位', '职级', '司龄']]
            for _, r in hp_df.iterrows():
                hp_data.append([str(r['姓名']), str(r['部门'])[:10], str(r['岗位'])[:10], str(r['职级']), f'{r["司龄"]:.1f}年'])
            t_hp = Table(hp_data, colWidths=[2.5*cm, 4*cm, 4*cm, 2*cm, 2.5*cm])
            t_hp.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#27AE60')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), chinese_font),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            story.append(t_hp)
        
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
    st.markdown("<h1 class='main-header'>产研团队全量人才盘点系统</h1>", unsafe_allow_html=True)
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
    | 📄 报告导出 | 导出PDF/Excel格式报告 |
    
    ### 📝 使用步骤
    1. 点击左侧 **"数据上传"** 上传员工信息表
    2. 点击 **"盘点分析"** 执行盘点
    3. 点击 **"报告导出"** 下载PDF/Excel报告
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
                total = len(result_df)
                st.session_state.stats = {
                    'total': total,
                    'departments': result_df['部门'].nunique() if '部门' in result_df.columns else 0,
                    'positions': result_df['职位'].nunique() if '职位' in result_df.columns else 0,
                    'full_match': len(result_df[result_df['匹配度'] == '完全匹配']),
                    'basic_match': len(result_df[result_df['匹配度'] == '基本匹配']),
                    'no_match': len(result_df[result_df['匹配度'] == '不匹配']),
                    'full_match_rate': len(result_df[result_df['匹配度'] == '完全匹配']) / total * 100,
                    'basic_match_rate': len(result_df[result_df['匹配度'] == '基本匹配']) / total * 100,
                    'no_match_rate': len(result_df[result_df['匹配度'] == '不匹配']) / total * 100,
                    'high_potential': int(result_df['高潜人才'].sum()),
                    'stable': int(result_df['稳定人员'].sum()),
                    'core_backbone': int(result_df['核心骨干'].sum()),
                    'attention': int(result_df['待关注'].sum()),
                    'optimization': int(result_df['待优化'].sum()),
                    'risk': int(result_df['离职风险'].sum()),
                    'high_potential_rate': result_df['高潜人才'].sum() / total * 100,
                    'stable_rate': result_df['稳定人员'].sum() / total * 100,
                    'core_backbone_rate': result_df['核心骨干'].sum() / total * 100,
                    'attention_rate': result_df['待关注'].sum() / total * 100,
                    'optimization_rate': result_df['待优化'].sum() / total * 100,
                    'risk_rate': result_df['离职风险'].sum() / total * 100,
                    'a_class': len(result_df[result_df['岗位分类'] == 'A类']),
                    'm_class': len(result_df[result_df['岗位分类'] == 'M类']),
                    'e_class': len(result_df[result_df['岗位分类'] == 'E类']),
                }
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
                st.bar_chart(layer_counts)
            with col2:
                st.markdown("#### 人岗匹配度分布")
                match_counts = df['匹配度'].value_counts()
                st.bar_chart(match_counts)

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
            st.markdown("#### 📊 Excel报告")
            st.markdown("包含完整盘点数据的Excel文件")
            if st.button("导出Excel", type="primary", key="excel_btn"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    overview = pd.DataFrame([
                        ['产研团队全量人才盘点报告'],
                        [f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}'],
                        [],
                        ['一、盘点概览'],
                        ['总人数', f'{stats["total"]}人'],
                        ['完全匹配', f'{stats["full_match"]}人 ({stats["full_match_rate"]:.1f}%)'],
                        ['待优化', f'{stats["optimization"]}人'],
                        [],
                        ['二、人才分层统计'],
                        ['高潜人才', f'{stats["high_potential"]}人'],
                        ['核心骨干', f'{stats["core_backbone"]}人'],
                        ['待优化', f'{stats["optimization"]}人'],
                    ])
                    overview.to_excel(writer, sheet_name='概览', index=False, header=False)
                    df.to_excel(writer, sheet_name='完整数据', index=False)
                output.seek(0)
                st.download_button(
                    "📥 下载 Excel 报告", 
                    data=output, 
                    file_name=f"人才盘点报告_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col2:
            st.markdown("#### 📄 PDF报告（汇报版）")
            st.markdown("适合汇报场景的PDF格式报告")
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
