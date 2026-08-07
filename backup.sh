#!/usr/bin/env bash
# 备份脚本：每次改动前运行，把当前 py 代码按时间戳备份到 backups/
# 用法: ./backup.sh [备注]
set -e

cd "$(dirname "$0")"
STAMP=$(date +%Y%m%d_%H%M%S)
NOTE="${1:-}"
DIR="backups/${STAMP}"
mkdir -p "$DIR"

# 备份所有源码与配置（backup.sh 自身所在目录的 py 源码 + tools + 公司画像）
cp -v spider_newest.py zbgg_crawler.py config.json "$DIR/" 2>/dev/null || true
cp -rv tools/*.py "$DIR/" 2>/dev/null || true
cp -rv company_profiles "$DIR/" 2>/dev/null || true

# 备注写入备份目录，便于追溯
if [ -n "$NOTE" ]; then
    echo "$NOTE" > "$DIR/NOTE.txt"
fi

# 归档为 zip（便于下载/解压）
if command -v zip >/dev/null 2>&1; then
    (cd backups && zip -q "${STAMP}.zip" "${STAMP}"/*)
fi

echo "备份完成: $DIR"
