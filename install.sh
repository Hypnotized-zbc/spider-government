#!/usr/bin/env bash
# 安装依赖（Windows + WSL2 环境）
set -e
cd "$(dirname "$0")"

echo "==> 创建/使用虚拟环境 venv"
if [ ! -d "venv" ]; then
  python3 -m venv venv
fi
source venv/bin/activate

echo "==> 安装依赖"
pip install -r requirements.txt

echo "==> 检查 Windows Edge（无头渲染依赖）"
if [ ! -f "/mnt/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe" ] \
   && [ ! -f "/mnt/c/Program Files/Microsoft/Edge/Application/msedge.exe" ]; then
  echo "警告: 未在标准路径找到 Microsoft Edge。程序需要 Windows 版 Edge 渲染详情页/下载附件。"
else
  echo "Edge 存在"
fi

echo "完成。运行: python3 zbgg_crawler.py"
