import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
import os

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    PDF_AVAILABLE = True
    
    CHINESE_FONT = None
    font_paths = [
        '/System/Library/Fonts/STHeiti Medium.ttc',
        '/System/Library/Fonts/STHeiti Light.ttc',
        '/System/Library/Fonts/Hiragino Sans GB.ttc',
        '/System/Library/Fonts/Supplemental/Songti.ttc',
        'C:/Windows/Fonts/simhei.ttf',
        'C:/Windows/Fonts/msyh.ttc',
    ]
    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont('ChineseFont', font_path, subfontIndex=0))
                CHINESE_FONT = 'ChineseFont'
                break
            except:
                continue
except ImportError:
    PDF_AVAILABLE = False
    CHINESE_FONT = None

class ReportGenerator:
    """报告生成器"""
    
    def __init__(self, df, stats, ai_analyzer=None):
        self.df = df
        self.stats = stats
        self.ai_analyzer = ai_analyzer
    
    def generate_excel(self):
        """生成Excel报告"""
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            overview_data = [
                ['产研团队全量人才盘点报告'],
                [f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}'],
                [],
                ['一、盘点概览'],
                ['盘点总人数', f'{self.stats["total"]}人'],
                ['部门覆盖', f'{self.stats["departments"]}个'],
                ['岗位类型', f'{self.stats["positions"]}种'],
                [],
                ['二、人才分层统计'],
                ['高潜人才', f'{self.stats["high_potential"]}人', f'{self.stats["high_potential_rate"]:.1f}%'],
                ['稳定人员', f'{self.stats["stable"]}人', f'{self.stats["stable_rate"]:.1f}%'],
                ['核心骨干', f'{self.stats["core_backbone"]}人', f'{self.stats["core_backbone_rate"]:.1f}%'],
                ['待关注', f'{self.stats["attention"]}人', f'{self.stats["attention_rate"]:.1f}%'],
                ['待优化', f'{self.stats["optimization"]}人', f'{self.stats["optimization_rate"]:.1f}%'],
                ['离职风险', f'{self.stats["risk"]}人', f'{self.stats["risk_rate"]:.1f}%'],
            ]
            pd.DataFrame(overview_data).to_excel(writer, sheet_name='盘点概览', index=False, header=False)
            
            if '部门' in self.df.columns:
                dept_stats = self.df.groupby('部门').agg({
                    '工号': 'count',
                    '高潜人才': 'sum',
                    '核心骨干': 'sum',
                    '待优化': 'sum'
                }).reset_index()
                dept_stats.columns = ['部门', '人数', '高潜人才', '核心骨干', '待优化']
                dept_stats.to_excel(writer, sheet_name='部门统计', index=False)
            
            talent_cols = ['工号', '姓名', '部门', '岗位', '职级', '岗位分类', 
                          '学历档位', '绩效档位', '匹配度', '司龄', '人才分层']
            available_cols = [col for col in talent_cols if col in self.df.columns]
            self.df[available_cols].to_excel(writer, sheet_name='完整数据', index=False)
            
            if '高潜人才' in self.df.columns:
                hp_df = self.df[self.df['高潜人才']][available_cols]
                hp_df.to_excel(writer, sheet_name='高潜人才', index=False)
            
            if '核心骨干' in self.df.columns:
                cb_df = self.df[self.df['核心骨干']][available_cols]
                cb_df.to_excel(writer, sheet_name='核心骨干', index=False)
            
            if '待优化' in self.df.columns:
                opt_df = self.df[self.df['待优化']][available_cols]
                opt_df.to_excel(writer, sheet_name='待优化人员', index=False)
            
            if '离职风险' in self.df.columns:
                risk_df = self.df[self.df['离职风险']][available_cols]
                risk_df.to_excel(writer, sheet_name='离职风险', index=False)
        
        output.seek(0)
        return output
    
    def generate_pdf(self):
        """生成PDF报告"""
        if not PDF_AVAILABLE or not CHINESE_FONT:
            return None
        
        output = BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4,
                               rightMargin=1.5*cm, leftMargin=1.5*cm,
                               topMargin=1.5*cm, bottomMargin=1.5*cm)
        
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontName=CHINESE_FONT,
            fontSize=22,
            textColor=colors.HexColor('#2C3E50'),
            spaceAfter=20,
            alignment=1
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontName=CHINESE_FONT,
            fontSize=14,
            textColor=colors.HexColor('#2C3E50'),
            spaceAfter=10,
            spaceBefore=15
        )
        
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontName=CHINESE_FONT,
            fontSize=10,
            leading=14
        )
        
        story = []
        
        story.append(Paragraph('产研团队全量人才盘点报告', title_style))
        story.append(Paragraph(f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}', normal_style))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('一、盘点概览', heading_style))
        
        overview_data = [
            ['指标', '数据'],
            ['盘点总人数', f'{self.stats["total"]}人'],
            ['部门覆盖', f'{self.stats["departments"]}个'],
            ['岗位类型', f'{self.stats["positions"]}种'],
            ['完全匹配', f'{self.stats["full_match"]}人 ({self.stats["full_match_rate"]:.1f}%)'],
            ['高潜人才', f'{self.stats["high_potential"]}人 ({self.stats["high_potential_rate"]:.1f}%)'],
            ['核心骨干', f'{self.stats["core_backbone"]}人 ({self.stats["core_backbone_rate"]:.1f}%)'],
            ['待优化', f'{self.stats["optimization"]}人 ({self.stats["optimization_rate"]:.1f}%)'],
            ['离职风险', f'{self.stats["risk"]}人 ({self.stats["risk_rate"]:.1f}%)'],
        ]
        
        table = Table(overview_data, colWidths=[6*cm, 10*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498DB')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ]))
        story.append(table)
        
        doc.build(story)
        output.seek(0)
        return output
    
    def generate_ai_enhanced_report(self):
        """生成AI增强报告"""
        if not self.ai_analyzer:
            return self.generate_pdf()
        
        if not PDF_AVAILABLE or not CHINESE_FONT:
            return None
        
        output = BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4,
                               rightMargin=1.5*cm, leftMargin=1.5*cm,
                               topMargin=1.5*cm, bottomMargin=1.5*cm)
        
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'CustomTitle', parent=styles['Heading1'],
            fontName=CHINESE_FONT, fontSize=22,
            textColor=colors.HexColor('#2C3E50'),
            spaceAfter=20, alignment=1
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading', parent=styles['Heading2'],
            fontName=CHINESE_FONT, fontSize=14,
            textColor=colors.HexColor('#2C3E50'),
            spaceAfter=10, spaceBefore=15
        )
        
        normal_style = ParagraphStyle(
            'CustomNormal', parent=styles['Normal'],
            fontName=CHINESE_FONT, fontSize=10, leading=14
        )
        
        story = []
        
        story.append(Paragraph('产研团队全量人才盘点报告', title_style))
        story.append(Paragraph('（AI增强版）', title_style))
        story.append(Paragraph(f'盘点日期：{datetime.now().strftime("%Y年%m月%d日")}', normal_style))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('一、盘点概览', heading_style))
        overview_data = [
            ['指标', '数据'],
            ['盘点总人数', f'{self.stats["total"]}人'],
            ['完全匹配率', f'{self.stats["full_match_rate"]:.1f}%'],
            ['高潜人才', f'{self.stats["high_potential"]}人'],
            ['核心骨干', f'{self.stats["core_backbone"]}人'],
            ['待优化', f'{self.stats["optimization"]}人'],
            ['离职风险', f'{self.stats["risk"]}人'],
        ]
        table = Table(overview_data, colWidths=[6*cm, 10*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498DB')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ]))
        story.append(table)
        story.append(Spacer(1, 20))
        
        story.append(Paragraph('二、AI智能分析', heading_style))
        try:
            ai_analysis = self.ai_analyzer.analyze_overall(self.df, self.stats)
            for line in ai_analysis.split('\n'):
                if line.strip():
                    story.append(Paragraph(line, normal_style))
                    story.append(Spacer(1, 5))
        except Exception as e:
            story.append(Paragraph(f'AI分析生成失败：{str(e)}', normal_style))
        
        story.append(PageBreak())
        
        story.append(Paragraph('三、AI优化建议', heading_style))
        try:
            suggestions = self.ai_analyzer.generate_suggestions(self.df, self.stats)
            for line in suggestions.split('\n'):
                if line.strip():
                    story.append(Paragraph(line, normal_style))
                    story.append(Spacer(1, 5))
        except Exception as e:
            story.append(Paragraph(f'建议生成失败：{str(e)}', normal_style))
        
        doc.build(story)
        output.seek(0)
        return output
