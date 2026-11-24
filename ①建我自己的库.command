#!/bin/bash
cd "$(dirname "$0")" || exit 1
echo "========== 第一步：建我自己的库（数据只在这台电脑）=========="
if ! command -v python3 >/dev/null 2>&1; then
  echo "没找到 python3。请到 https://www.python.org/downloads/ 下载安装后再双击本文件。"
  read -r -p "按回车关闭…" _; exit 1
fi
read -r -p "请输入你的名字（例如 张三），然后回车：" NAME
if [ -z "$NAME" ]; then echo "名字不能为空。"; read -r -p "按回车关闭…" _; exit 1; fi
python3 scripts/init_person.py --name "$NAME" || { read -r -p "按回车关闭…" _; exit 1; }
echo "data/persons/$NAME/redpacket.db" > config/my_db.txt
echo "已记住你的库：data/persons/$NAME/redpacket.db"
echo "接下来双击 ②导入我的表格并核对.command"
read -r -p "按回车关闭…" _
