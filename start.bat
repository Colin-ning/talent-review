@echo off
echo ========================================
echo 产研团队人才盘点工具 - 启动服务
echo ========================================
echo.

if not exist "venv" (
    echo 首次运行，正在创建虚拟环境...
    python -m venv venv
)

call venv\Scripts\activate

echo 正在安装依赖...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

echo.
echo 启动Streamlit服务...
echo 请在浏览器中访问: http://localhost:8501
echo.

streamlit run app.py
pause
