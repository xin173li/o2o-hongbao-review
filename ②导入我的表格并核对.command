#!/bin/bash
cd "$(dirname "$0")" || exit 1
echo "========== 第二步：导入我自己的底表并自动核对 =========="
if [ ! -f config/my_db.txt ]; then echo "还没建库，请先双击 ①建我自己的库.command"; read -r -p "按回车关闭…" _; exit 1; fi
DB=$(cat config/my_db.txt)
echo "我的库：$DB"
python3 scripts/import_data.py --db "$DB" || { read -r -p "按回车关闭…" _; exit 1; }
python3 scripts/run_audit.py --db "$DB" || { read -r -p "按回车关闭…" _; exit 1; }
echo "报告在 data/out/ 里，最新一份叫 最新核对报告.xlsx"
read -r -p "按回车关闭…" _
