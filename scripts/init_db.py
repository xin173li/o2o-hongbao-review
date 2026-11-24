"""初始化数据库：建表 + 维度基础数据。"""
from __future__ import annotations

import argparse

from common import ROOT, connect, db_file, ensure_schema, log, load_settings, use_db

CHANNELS = [("京东", "京东"), ("天猫", "天猫"), ("零售云", "零售云")]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reset", action="store_true", help="删除已有数据库后重建")
    parser.add_argument("--db", help="指定数据库文件（默认 data/redpacket.db）")
    args = parser.parse_args()
    use_db(args.db)

    db = connect()
    if args.reset:
        db.close()
        from common import path_of
        p = path_of("db")
        for suffix in ("", "-wal", "-shm"):
            f = p.with_name(p.name + suffix) if suffix else p
            if f.exists():
                f.unlink()
        log(f"已删除数据库 {p.name}")
        db = connect()

    ensure_schema(db)
    with db:
        for code, name in CHANNELS:
            db.execute(
                "INSERT INTO dim_channel (channel_code, channel_name) VALUES (?, ?) "
                "ON CONFLICT(channel_code) DO UPDATE SET channel_name = excluded.channel_name",
                (code, name),
            )

    cfg = load_settings()["audit"]
    counts = {
        "policy": db.execute("SELECT COUNT(*) FROM policy").fetchone()[0],
        "order": db.execute("SELECT COUNT(*) FROM fact_order").fetchone()[0],
        "issue": db.execute("SELECT COUNT(*) FROM fact_issue").fetchone()[0],
        "opening": db.execute("SELECT COUNT(*) FROM opening_balance").fetchone()[0],
    }
    log(f"当前数据库：{db_file().relative_to(ROOT)}")
    log(f"数据库就绪｜基准月 {cfg['baseline_period']}｜判定起始 {cfg['judge_from_period']}")
    log(f"现有数据：{counts}")
    db.close()


if __name__ == "__main__":
    main()
