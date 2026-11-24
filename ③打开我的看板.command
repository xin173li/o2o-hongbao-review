#!/bin/bash
cd "$(dirname "$0")" || exit 1
if [ ! -f config/my_db.txt ]; then echo "还没建库，请先双击 ①建我自己的库.command"; read -r -p "按回车关闭…" _; exit 1; fi
DB=$(cat config/my_db.txt)
echo "正在打开看板：http://127.0.0.1:8899/"
python3 scripts/serve_dashboard.py --db "$DB" --port 8899 &
sleep 3
open "http://127.0.0.1:8899/" 2>/dev/null || echo "请在浏览器打开 http://127.0.0.1:8899/"
wait
