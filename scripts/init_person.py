"""每人一套自己的数据库：创建个人库并落一份使用说明。数据不出本机。

用法：
  python scripts/init_person.py --name 张三
  python scripts/init_person.py --name 张三 --dir data/persons
"""
from __future__ import annotations

import argparse
from datetime import datetime

from common import ROOT, connect, ensure_schema, load_settings, log, use_db

GUIDE = """O2O 红包核销 · 个人数据库使用说明

姓名：{name}
数据库：{db}
创建时间：{when}

这套库里只有你自己的数据，别人不会写进来，你也不会写到别人的库里。

第 1 步 每次打开命令行，先指向自己的库（Windows PowerShell 用 $env:）

  export REDPACKET_DB={db}

第 2 步 导入自己的底表（把 Excel 放进 data/inbox/，或用网页上传）

  python scripts/import_data.py

第 3 步 核对并出报告

  python scripts/run_audit.py

第 4 步 打开自己的看板

  python scripts/serve_dashboard.py --port 8899
  # 浏览器访问 http://127.0.0.1:8899/ ，页面右上角会显示当前库文件名

也可以用参数临时指定，不必设环境变量：

  python scripts/run_audit.py --db {db}
  python scripts/serve_dashboard.py --db {db} --port 8899

口径（判定起始月、期初基准月、金额/数量容忍度）在 config/settings.json，
改完直接重跑核对即可，不需要改代码。
"""


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--name", required=True, help="使用人姓名，例如 张三")
    parser.add_argument("--dir", default="data/persons", help="个人库所在目录（默认 data/persons）")
    parser.add_argument("--reset", action="store_true", help="已存在则清空重建")
    args = parser.parse_args()

    person_dir = (ROOT / args.dir / args.name).resolve()
    person_dir.mkdir(parents=True, exist_ok=True)
    db_path = person_dir / "redpacket.db"

    if db_path.exists() and args.reset:
        for suffix in ("", "-wal", "-shm"):
            f = db_path.with_name(db_path.name + suffix) if suffix else db_path
            if f.exists():
                f.unlink()

    use_db(str(db_path))
    conn = connect()
    ensure_schema(conn)
    counts = {
        "policy": conn.execute("SELECT COUNT(*) FROM policy").fetchone()[0],
        "order": conn.execute("SELECT COUNT(*) FROM fact_order").fetchone()[0],
        "issue": conn.execute("SELECT COUNT(*) FROM fact_issue").fetchone()[0],
        "opening": conn.execute("SELECT COUNT(*) FROM opening_balance").fetchone()[0],
    }
    conn.close()

    guide = person_dir / "使用说明.txt"
    guide.write_text(GUIDE.format(name=args.name, db=db_path.relative_to(ROOT),
                                 when=datetime.now().strftime("%Y-%m-%d %H:%M")), encoding="utf-8")

    cfg = load_settings()["audit"]
    log(f"已为「{args.name}」建好个人库：{db_path.relative_to(ROOT)}")
    log(f"现有数据：{counts}｜基准月 {cfg['baseline_period']}｜判定起始 {cfg['judge_from_period']}")
    log(f"使用说明：{guide.relative_to(ROOT)}")
    log(f"后续命令前面加：export REDPACKET_DB={db_path.relative_to(ROOT)}（Windows：$env:REDPACKET_DB=\"{db_path.relative_to(ROOT)}\"）")


if __name__ == "__main__":
    main()
