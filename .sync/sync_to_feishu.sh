#!/bin/bash
# Workspace 到飞书云文档同步脚本
# 运行时间: $(date)

WORKSPACE_DIR="/root/.openclaw/workspace"
LOG_FILE="/root/.openclaw/workspace/.sync/sync.log"
SYNC_STATE="/root/.openclaw/workspace/.sync/last_sync"

# 确保日志目录存在
mkdir -p $(dirname $LOG_FILE)

echo "=== 同步开始: $(date) ===" >> $LOG_FILE

# 统计文件
PPT_COUNT=$(find $WORKSPACE_DIR -maxdepth 1 -name "*.pptx" -o -name "*.ppt" | wc -l)
MD_COUNT=$(find $WORKSPACE_DIR -maxdepth 1 -name "*.md" | wc -l)
PY_COUNT=$(find $WORKSPACE_DIR -maxdepth 1 -name "*.py" | wc -l)
IMG_COUNT=$(find $WORKSPACE_DIR -maxdepth 1 -name "*.png" -o -name "*.jpg" | wc -l)

echo "文件统计:" >> $LOG_FILE
echo "  - PPT文件: $PPT_COUNT" >> $LOG_FILE
echo "  - Markdown文档: $MD_COUNT" >> $LOG_FILE
echo "  - Python脚本: $PY_COUNT" >> $LOG_FILE
echo "  - 图片: $IMG_COUNT" >> $LOG_FILE

# 记录同步时间
date +%s > $SYNC_STATE
echo "同步完成: $(date)" >> $LOG_FILE
echo "" >> $LOG_FILE
