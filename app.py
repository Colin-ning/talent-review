import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import os
import sys
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import base64

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from utils.talent_review import TalentReviewer
from utils.ai_analyzer import AIAnalyzer
from utils.report_generator import ReportGenerator
from utils.industry_data import IndustryDataFetcher

st.set_page_config(
    page_title="产研团队人才盘点系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
    }
    .stMetric > div {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

if 'reviewer' not in st.session_state:
    st.session_state.reviewer = None
if 'df' not in st.session_state:
    st.session_state.df = None
if 'review_complete' not in st.session_state:
    st.session_state.review_complete = False
if 'ai_analyzer' not in st.session_state:
    st.session_state.ai_analyzer = None

with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/000000/combo-chart.png", width=80)
    st.title("产研团队人才盘点系统")
    st.markdown("---")
    
    api_key = st.text_input("智谱AI API Key", type="password", help="请输入您的智谱AI API Key")
    if api_key:
        st.session_state.ai_analyzer = AIAnalyzer(api_key)
        st.success("✅ API Key 已配置")
    
    st.markdown("---")
    st.markdown("### 📋 功能导航")
    page = st.radio(
        "选择功能",
        ["🏠 首页概览", "📤 数据上传", "📊 盘点分析", "🤖 AI智能分析", "📈 行业对比", "💬 智能问答", "📄 报告导出"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    st.markdown(f"""
    <div style='text-align: center; color: #666; font-size: 0.8rem;'>
        版本：v2.0 Web版<br>
        更新日期：{datetime.now().strftime('%Y-%m-%d')}
    </div>
    """, unsafe_allow_html=True)

if page == "🏠 首页概览":
    st.markdown("<h1 class='main-header'>产研团队全量人才盘点系统</h1>", unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(label="👥 总人数", value="--", delta="待上传数据")
    with col2:
        st.metric(label="🏢 部门数", value="--", delta="待上传数据")
    with col3:
        st.metric(label="✅ 匹配率", value="--", delta="待分析")
    with col4:
        st.metric(label="⚠️ 风险人数", value="--", delta="待分析")
    
    st.markdown("---")
    
    st.markdown("""
    ### 🎯 系统功能
    
    本系统基于**智谱AI GLM-5大模型**，提供智能化的人才盘点分析服务：
    
    | 功能模块 | 说明 |
    |---------|------|
    | 📤 数据上传 | 上传员工基础信息表，自动解析数据 |
    | 📊 盘点分析 | 执行人才盘点，生成详细分析结果 |
    | 🤖 AI智能分析 | AI分析人才数据，提供个性化建议 |
    | 📈 行业对比 | 实时获取行业数据，进行对比分析 |
    | 💬 智能问答 | 自然语言交互，解答盘点相关问题 |
    | 📄 报告导出 | 导出PDF/Excel格式报告 |
    
    ### 📝 使用步骤
    
    1. **配置API**：在左侧输入智谱AI API Key
    2. **上传数据**：点击"数据上传"，上传员工信息表
    3. **执行盘点**：点击"盘点分析"，生成盘点结果
    4. **AI分析**：使用AI功能获取智能建议
    5. **导出报告**：下载PDF/Excel格式报告
    """)
    
    st.info("💡 提示：请先在左侧输入智谱AI API Key，然后上传员工数据开始使用")

elif page == "📤 数据上传":
    st.header("📤 数据上传")
    
    st.markdown("""
    ### 上传员工基础信息表
    
    请上传Excel格式的员工基础信息表，文件需包含以下字段：
    - 基本信息：工号、姓名、入职时间、岗位类型、部门、岗位、职位、职级
    - 学历信息：学历-学校类别、毕业学校、海外高校QS排名
    - 岗位分类：岗位分类（A类/M类/E类）
    - 绩效信息：2024上半年/年度、2025上半年/年度绩效结果
    """)
    
    uploaded_file = st.file_uploader("选择文件", type=['xlsx', 'xls'], help="支持.xlsx和.xls格式")
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ 文件上传成功！共读取 {len(df)} 条记录")
            
            st.markdown("### 📋 数据预览")
            st.dataframe(df.head(10), use_container_width=True)
            
            st.markdown("### 📊 数据统计")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("总人数", f"{len(df)}人")
            with col2:
                st.metric("部门数", f"{df['部门'].nunique()}个")
            with col3:
                st.metric("岗位数", f"{df['职位'].nunique()}种")
            
            required_fields = ['工号', '姓名', '入职时间', '部门', '岗位', '职级', '岗位分类']
            missing_fields = [f for f in required_fields if f not in df.columns]
            
            if missing_fields:
                st.error(f"❌ 缺少必填字段：{', '.join(missing_fields)}")
            else:
                if st.button("✅ 确认使用此数据", type="primary"):
                    st.session_state.df = df
                    st.session_state.reviewer = TalentReviewer(df)
                    st.session_state.review_complete = False
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
                st.session_state.reviewer.perform_review()
                st.session_state.review_complete = True
            st.success("✅ 盘点分析完成！")
            st.rerun()
        
        if st.session_state.review_complete:
            df = st.session_state.reviewer.get_result_df()
            stats = st.session_state.reviewer.get_statistics()
            
            st.markdown("---")
            st.markdown("### 📈 盘点结果概览")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("总人数", f"{stats['total']}人")
            with col2:
                st.metric("完全匹配", f"{stats['full_match']}人", f"{stats['full_match_rate']:.1f}%")
            with col3:
                st.metric("待优化", f"{stats['optimization']}人", f"{stats['optimization_rate']:.1f}%")
            with col4:
                st.metric("离职风险", f"{stats['risk']}人", f"{stats['risk_rate']:.1f}%")
            
            st.markdown("---")
            
            tab1, tab2, tab3, tab4 = st.tabs(["人才分层", "岗位分类", "职级分布", "匹配度分析"])
            
            with tab1:
                st.markdown("### 人才分层统计")
                talent_data = {
                    '分类': ['高潜人才', '稳定人员', '核心骨干', '待关注', '待优化', '离职风险'],
                    '人数': [stats['high_potential'], stats['stable'], stats['core_backbone'], 
                            stats['attention'], stats['optimization'], stats['risk']],
                    '占比': [f"{stats['high_potential']/stats['total']*100:.1f}%",
                            f"{stats['stable']/stats['total']*100:.1f}%",
                            f"{stats['core_backbone']/stats['total']*100:.1f}%",
                            f"{stats['attention']/stats['total']*100:.1f}%",
                            f"{stats['optimization']/stats['total']*100:.1f}%",
                            f"{stats['risk']/stats['total']*100:.1f}%"]
                }
                st.dataframe(pd.DataFrame(talent_data), use_container_width=True)
                
                fig = px.pie(df, names='人才分层', title='人才分层分布', hole=0.4)
                st.plotly_chart(fig, use_container_width=True)
            
            with tab2:
                st.markdown("### 岗位分类分布")
                class_stats = df['岗位分类'].value_counts().reset_index()
                class_stats.columns = ['岗位分类', '人数']
                fig = px.bar(class_stats, x='岗位分类', y='人数', color='岗位分类',
                            title='岗位分类分布', text='人数')
                st.plotly_chart(fig, use_container_width=True)
            
            with tab3:
                st.markdown("### 职级分布")
                level_stats = df['职级'].value_counts().reset_index()
                level_stats.columns = ['职级', '人数']
                fig = px.bar(level_stats, x='职级', y='人数', color='职级',
                            title='职级分布', text='人数')
                st.plotly_chart(fig, use_container_width=True)
            
            with tab4:
                st.markdown("### 人岗匹配度分析")
                match_stats = df['匹配度'].value_counts().reset_index()
                match_stats.columns = ['匹配度', '人数']
                fig = px.pie(match_stats, values='人数', names='匹配度', 
                            title='人岗匹配度分布', color='匹配度',
                            color_discrete_map={'完全匹配': '#2ECC71', '基本匹配': '#F39C12', '不匹配': '#E74C3C'})
                st.plotly_chart(fig, use_container_width=True)
                
                st.markdown("### 各部门匹配率")
                dept_match = df.groupby('部门').apply(
                    lambda x: (x['匹配度'] == '完全匹配').sum() / len(x) * 100
                ).reset_index()
                dept_match.columns = ['部门', '匹配率']
                dept_match['匹配率'] = dept_match['匹配率'].round(1)
                fig = px.bar(dept_match.sort_values('匹配率', ascending=False), 
                            x='部门', y='匹配率', color='匹配率',
                            title='各部门人岗匹配率(%)')
                st.plotly_chart(fig, use_container_width=True)

elif page == "🤖 AI智能分析":
    st.header("🤖 AI智能分析")
    
    if not st.session_state.review_complete:
        st.warning("⚠️ 请先完成盘点分析")
    elif st.session_state.ai_analyzer is None:
        st.warning("⚠️ 请先在左侧配置智谱AI API Key")
    else:
        st.markdown("### AI智能人才分析")
        
        df = st.session_state.reviewer.get_result_df()
        stats = st.session_state.reviewer.get_statistics()
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📊 整体分析")
            if st.button("生成整体分析报告", type="primary"):
                with st.spinner("AI正在分析中..."):
                    analysis = st.session_state.ai_analyzer.analyze_overall(df, stats)
                st.markdown("##### AI分析结果")
                st.markdown(analysis)
        
        with col2:
            st.markdown("#### 🎯 优化建议")
            if st.button("生成优化建议", type="primary"):
                with st.spinner("AI正在生成建议..."):
                    suggestions = st.session_state.ai_analyzer.generate_suggestions(df, stats)
                st.markdown("##### AI建议")
                st.markdown(suggestions)
        
        st.markdown("---")
        st.markdown("#### 👤 个人分析")
        
        employee = st.selectbox("选择员工", df['姓名'].tolist())
        if st.button("分析该员工"):
            with st.spinner("AI正在分析..."):
                employee_data = df[df['姓名'] == employee].iloc[0].to_dict()
                personal_analysis = st.session_state.ai_analyzer.analyze_employee(employee_data)
            st.markdown("##### 个人分析结果")
            st.markdown(personal_analysis)

elif page == "📈 行业对比":
    st.header("📈 行业数据对比")
    
    if not st.session_state.review_complete:
        st.warning("⚠️ 请先完成盘点分析")
    else:
        st.markdown("### 行业薪酬与人才数据对比")
        
        df = st.session_state.reviewer.get_result_df()
        
        industry_fetcher = IndustryDataFetcher()
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📊 行业薪酬水平")
            position = st.selectbox("选择职位", df['职位'].unique().tolist()[:10])
            if st.button("查询行业薪酬"):
                with st.spinner("正在获取行业数据..."):
                    salary_data = industry_fetcher.get_salary_data(position)
                st.markdown(f"**{position}** 行业薪酬数据：")
                st.dataframe(pd.DataFrame(salary_data), use_container_width=True)
        
        with col2:
            st.markdown("#### 📈 人才流动趋势")
            if st.button("查询人才趋势"):
                with st.spinner("正在获取趋势数据..."):
                    trend_data = industry_fetcher.get_talent_trend()
                st.markdown("**产研人才市场趋势**")
                st.dataframe(pd.DataFrame(trend_data), use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### 🏢 团队与行业对比")
        
        if st.button("生成对比分析"):
            with st.spinner("正在生成对比分析..."):
                comparison = industry_fetcher.compare_with_industry(df)
            st.markdown("##### 对比结果")
            st.dataframe(pd.DataFrame(comparison), use_container_width=True)

elif page == "💬 智能问答":
    st.header("💬 智能问答助手")
    
    if st.session_state.ai_analyzer is None:
        st.warning("⚠️ 请先在左侧配置智谱AI API Key")
    else:
        st.markdown("""
        ### 💡 智能问答
        
        您可以用自然语言提问关于人才盘点的问题，AI将为您解答。
        
        **示例问题：**
        - 如何判断一个员工是否属于高潜人才？
        - 待优化员工应该如何处理？
        - 如何降低核心骨干的离职风险？
        - 人才盘点的主要目的是什么？
        """)
        
        if 'chat_history' not in st.session_state:
            st.session_state.chat_history = []
        
        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
        
        if prompt := st.chat_input("输入您的问题..."):
            st.session_state.chat_history.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)
            
            with st.chat_message("assistant"):
                with st.spinner("思考中..."):
                    context = ""
                    if st.session_state.review_complete:
                        stats = st.session_state.reviewer.get_statistics()
                        context = f"当前盘点数据：{stats}"
                    
                    response = st.session_state.ai_analyzer.chat(prompt, context)
                st.markdown(response)
                st.session_state.chat_history.append({"role": "assistant", "content": response})
        
        if st.button("清空对话"):
            st.session_state.chat_history = []
            st.rerun()

elif page == "📄 报告导出":
    st.header("📄 报告导出")
    
    if not st.session_state.review_complete:
        st.warning("⚠️ 请先完成盘点分析")
    else:
        df = st.session_state.reviewer.get_result_df()
        stats = st.session_state.reviewer.get_statistics()
        
        st.markdown("### 选择导出格式")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📊 Excel报告")
            st.markdown("包含16个工作表的详细盘点报告")
            if st.button("导出Excel", type="primary"):
                with st.spinner("正在生成Excel报告..."):
                    generator = ReportGenerator(df, stats)
                    excel_buffer = generator.generate_excel()
                    
                    st.download_button(
                        label="下载 Excel 报告",
                        data=excel_buffer,
                        file_name=f"人才盘点报告_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        with col2:
            st.markdown("#### 📄 PDF报告")
            st.markdown("适合汇报场景的PDF格式报告")
            if st.button("导出PDF", type="primary"):
                with st.spinner("正在生成PDF报告..."):
                    generator = ReportGenerator(df, stats)
                    pdf_buffer = generator.generate_pdf()
                    
                    st.download_button(
                        label="下载 PDF 报告",
                        data=pdf_buffer,
                        file_name=f"人才盘点报告_汇报版_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )
        
        st.markdown("---")
        st.markdown("#### 🤖 AI增强报告")
        
        if st.session_state.ai_analyzer is None:
            st.warning("⚠️ 请先配置智谱AI API Key以使用AI增强报告功能")
        else:
            st.markdown("包含AI智能分析和建议的增强版报告")
            if st.button("生成AI增强报告", type="primary"):
                with st.spinner("AI正在生成增强报告..."):
                    generator = ReportGenerator(df, stats, st.session_state.ai_analyzer)
                    enhanced_pdf = generator.generate_ai_enhanced_report()
                    
                    st.download_button(
                        label="下载 AI增强报告",
                        data=enhanced_pdf,
                        file_name=f"人才盘点报告_AI增强版_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )
