# 产研团队人才盘点系统 - Web版

## 系统简介

基于 **Streamlit** 和 **智谱AI GLM-5** 的智能化人才盘点系统，提供以下核心功能：

- 📊 **数据上传**：支持Excel格式员工数据上传
- 📈 **盘点分析**：自动执行人才盘点分析
- 🤖 **AI智能分析**：AI分析人才数据，提供个性化建议
- 📊 **行业对比**：实时获取行业数据，进行对比分析
- 💬 **智能问答**：自然语言交互，解答盘点相关问题
- 📄 **报告导出**：导出PDF/Excel格式报告

## 技术栈

- **前端框架**：Streamlit
- **AI模型**：智谱AI GLM-5
- **数据处理**：Pandas, NumPy
- **可视化**：Plotly, Matplotlib
- **报告生成**：ReportLab, OpenPyXL

## 快速开始

### 方式一：本地运行

1. **安装依赖**
```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

2. **启动服务**
```bash
streamlit run app.py
```

3. **访问应用**
打开浏览器访问：http://localhost:8501

### 方式二：使用启动脚本

**Mac/Linux:**
```bash
chmod +x start.sh
./start.sh
```

**Windows:**
双击运行 `start.bat`

### 方式三：Docker部署

```bash
docker-compose up -d
```

## 配置说明

### 智谱AI API配置

1. 访问 [智谱AI开放平台](https://open.bigmodel.cn/)
2. 注册账号并创建API Key
3. 在应用侧边栏输入API Key

或配置环境变量：
```bash
export ZHIPUAI_API_KEY="your_api_key_here"
```

## 功能模块

### 1. 数据上传
- 支持Excel格式(.xlsx, .xls)
- 自动解析必填字段
- 数据预览和统计

### 2. 盘点分析
- 学历档位判定
- 绩效档位判定
- 人岗匹配度分析
- 人才分层盘点
- 风险预警

### 3. AI智能分析
- 整体分析报告
- 优化建议生成
- 个人发展分析

### 4. 行业对比
- 行业薪酬数据
- 人才流动趋势
- 团队与行业对比

### 5. 智能问答
- 自然语言交互
- 专业问题解答
- 盘点知识库

### 6. 报告导出
- Excel详细报告
- PDF汇报报告
- AI增强报告

## 云服务器部署

### 使用Docker部署（推荐）

1. **安装Docker和Docker Compose**
```bash
curl -fsSL https://get.docker.com | sh
sudo usermod -aG docker $USER
```

2. **克隆项目**
```bash
git clone <your-repo-url>
cd 产研团队人才盘点工具_Web版
```

3. **启动服务**
```bash
docker-compose up -d
```

4. **配置域名和SSL（可选）**

使用Nginx反向代理：
```nginx
server {
    listen 80;
    server_name your-domain.com;
    
    location / {
        proxy_pass http://localhost:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
    }
}
```

### 使用Streamlit Cloud部署

1. 将代码推送到GitHub
2. 访问 [Streamlit Cloud](https://streamlit.io/cloud)
3. 连接GitHub仓库并部署

## 目录结构

```
产研团队人才盘点工具_Web版/
├── app.py                 # 主应用
├── requirements.txt       # 依赖列表
├── Dockerfile            # Docker配置
├── docker-compose.yml    # Docker Compose配置
├── start.sh              # Mac启动脚本
├── start.bat             # Windows启动脚本
├── utils/                # 工具模块
│   ├── __init__.py
│   ├── talent_review.py  # 盘点核心逻辑
│   ├── ai_analyzer.py    # AI分析模块
│   ├── industry_data.py  # 行业数据模块
│   └── report_generator.py # 报告生成模块
└── pages/                # 多页面（可选）
```

## 盘点规则说明

### 学历档位定义
- **一档**：C9/985院校、QS≤300
- **二档**：211/双一流院校、QS≤500
- **三档**：其他本科及以上

### 绩效档位定义
- **一档**：S/A占比>50%，年度S/A，其余B+以上
- **二档**：全部B+以上
- **三档**：至多1次B，其余B+以上

### 人岗匹配标准
- **完全匹配**：学历+绩效双优
- **基本匹配**：满足门槛，无C级绩效
- **不匹配**：不满足门槛或有C级绩效

### 人才分层标准
- **高潜人才**：司龄≤1年 + 近2期全A/S + 职级≤L6
- **稳定人员**：司龄≥0.5年 + 4期全B+以上
- **核心骨干**：A/M类 + 绩效优 + L6以上 + 完全匹配
- **待关注**：绩效波动或基本匹配
- **待优化**：有C级绩效或不匹配

## 注意事项

1. **数据安全**：员工数据仅用于本地分析
2. **API密钥**：请妥善保管智谱AI API Key
3. **浏览器兼容**：推荐使用Chrome/Firefox
4. **网络要求**：AI功能需要网络连接

## 版本信息

- **版本号**：v2.0 Web版
- **更新日期**：2026年03月26日
- **开发环境**：Python 3.9

## 技术支持

如有问题，请联系工具开发者。
