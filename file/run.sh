#!/bin/bash

echo "示例Excel数据生成器"
echo "==================="
echo

echo "检查Python环境..."
python3 --version
if [ $? -ne 0 ]; then
    echo "错误: 未找到Python环境，请先安装Python"
    exit 1
fi

echo
echo "安装依赖包..."
pip3 install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "警告: 依赖包安装可能有问题，但继续执行..."
fi

echo
echo "开始生成Excel数据..."
python3 generate_excel_sample.py

echo
echo "生成完成！"
echo "请查看当前目录下的 sample_large_excel_data.xlsx 文件"
echo 