#!/usr/bin/env bash
# 一键跑全流程：导入 data/inbox 下所有底表 -> 核对 -> 出报告
set -e
cd "$(dirname "$0")"
PY=${PYTHON:-python3}

echo "== 1/3 初始化数据库 =="
$PY scripts/init_db.py

echo "== 2/3 导入底表 =="
$PY scripts/import_data.py

echo "== 3/3 核对并出报告 =="
$PY scripts/run_audit.py

echo "完成。报告在 data/out/"
