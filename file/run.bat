@echo off
echo 示例Excel数据生成器
echo ===================
echo.

echo 检查Python环境...
python --version
if errorlevel 1 (
    echo 错误: 未找到Python环境，请先安装Python
    pause
    exit /b 1
)

echo.
echo 安装依赖包...
pip install -r requirements.txt
if errorlevel 1 (
    echo 警告: 依赖包安装可能有问题，但继续执行...
)

echo.
echo 开始生成Excel数据...
python generate_excel_sample.py

echo.
echo 生成完成！
echo 请查看当前目录下的 sample_large_excel_data.xlsx 文件
echo.
pause 