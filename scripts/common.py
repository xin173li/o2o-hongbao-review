"""公共工具：配置、路径、字段归一化、数据库连接。"""
from __future__ import annotations

import json
import os
import re
import sqlite3
from datetime import datetime, date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SETTINGS_PATH = ROOT / "config" / "settings.json"

DEFAULT_SETTINGS = {
    "paths": {"db": "data/redpacket.db", "inbox": "data/inbox",
              "archive": "data/archive", "out": "data/out", "web": "web"},
    "audit": {"baseline_period": "2025-12", "judge_from_period": "2026-01",
              "amount_tolerance": 0.5, "qty_tolerance": 0.0001},
    "import": {"archive_after_import": True},
    "source_rules": {"policy_validity": "month", "model_ignore_case_and_space": True,
                     "mixed_type_within_row": True},
}


def load_settings() -> dict:
    cfg = json.loads(json.dumps(DEFAULT_SETTINGS))
    if SETTINGS_PATH.exists():
        user = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        for section, value in user.items():
            if isinstance(value, dict):
                cfg.setdefault(section, {}).update(value)
            else:
                cfg[section] = value
    return cfg


def db_file() -> Path:
    """当前使用的数据库文件。可用环境变量 REDPACKET_DB 指定自己的一套库。"""
    env = os.environ.get("REDPACKET_DB")
    if env:
        p = Path(env).expanduser()
        return p if p.is_absolute() else (ROOT / p)
    return ROOT / load_settings()["paths"]["db"]


def use_db(value: str | None) -> None:
    """脚本参数 --db 的落地方式：设置环境变量，后续所有连接都指向这个库。"""
    if value:
        os.environ["REDPACKET_DB"] = str(value)


def path_of(key: str) -> Path:
    """返回配置中的路径（相对项目根目录），并确保父目录存在。"""
    if key == "db":
        p = db_file()
        p.parent.mkdir(parents=True, exist_ok=True)
        return p
    cfg = load_settings()
    p = ROOT / cfg["paths"][key]
    p.mkdir(parents=True, exist_ok=True)
    return p


# ---------- 期间 ----------
EXCEL_EPOCH = date(1899, 12, 30)


def norm_period(value, default_year: int | None = None) -> str | None:
    """把各种月份写法归一为 YYYY-MM。
    支持：2026-08 / 2026/8 / 202608 / datetime / Excel 序列号 / 8（需 default_year）
    """
    if value is None or value == "":
        return None
    if isinstance(value, (datetime, date)):
        return f"{value.year:04d}-{value.month:02d}"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel 序列号（> 1000 才当日期，否则当月份数字）
        if value > 1000:
            dt = EXCEL_EPOCH.fromordinal(EXCEL_EPOCH.toordinal() + int(value))
            return f"{dt.year:04d}-{dt.month:02d}"
        month = int(value)
        if default_year and 1 <= month <= 12:
            return f"{default_year:04d}-{month:02d}"
        return None
    s = str(value).strip()
    s = s.replace("年", "-").replace("月", "").replace("/", "-").replace(".", "-")
    m = re.match(r"^(\d{4})-(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    m = re.match(r"^(\d{4})(\d{2})$", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.match(r"^(\d{1,2})$", s)
    if m and default_year and 1 <= int(m.group(1)) <= 12:
        return f"{default_year:04d}-{int(m.group(1)):02d}"
    try:
        dt = datetime.fromisoformat(s)
        return f"{dt.year:04d}-{dt.month:02d}"
    except Exception:
        return None


def add_months(period: str, n: int) -> str:
    y, m = int(period[:4]), int(period[5:7])
    m += n
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return f"{y:04d}-{m:02d}"


def period_range(start: str, end: str) -> list[str]:
    out, cur = [], start
    while cur <= end:
        out.append(cur)
        cur = add_months(cur, 1)
        if len(out) > 240:
            break
    return out


# ---------- 文本 / 数值 ----------
def norm_model(value) -> str | None:
    """型号归一：去空格、转大写（不区分大小写与空格）。"""
    if value is None:
        return None
    s = str(value).replace(" ", "").replace("\u3000", "").strip().upper()
    return s or None


def norm_text(value) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    return s or None


def to_num(value) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    s = str(value).strip().replace(",", "").replace("元", "").replace("￥", "")
    s = s.replace("（", "(").replace("）", ")")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def red_type_of(value) -> list[str]:
    """红包类型归一：一行可能同时包含两类（混写），返回列表。"""
    if value is None or str(value).strip() == "":
        return []
    s = str(value)
    out = []
    if "零售" in s:
        out.append("retail")
    if "下单" in s or "预售" in s:
        out.append("order")
    if not out:
        # 常见英文/简写
        low = s.lower()
        if "retail" in low:
            out.append("retail")
        if "order" in low:
            out.append("order")
    return list(dict.fromkeys(out))


# ---------- 数据库 ----------
def connect() -> sqlite3.Connection:
    db = db_file()
    db.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def ensure_schema(conn: sqlite3.Connection) -> None:
    sql = (ROOT / "db" / "schema.sql").read_text(encoding="utf-8")
    conn.executescript(sql)
    # 兼容现有系统导出的字段：老库自动补列，不丢数据
    for table, col, ctype in (("fact_order", "order_date", "TEXT"), ("fact_order", "category", "TEXT"),
                              ("fact_order", "region", "TEXT"), ("fact_issue", "category", "TEXT"),
                              ("fact_issue", "region", "TEXT")):
        existing = [r[1] for r in conn.execute(f"PRAGMA table_info({table})")]
        if col not in existing:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ctype}")
    conn.commit()


def log(msg: str, level: str = "INFO") -> None:
    print(f"[{level}] {msg}")


_ALIAS = {
    "channel": ["渠道", "渠道名称", "平台"],
    "store_code": ["门店编码", "门店代码", "客户编码", "门店ID"],
    "store_name": ["门店名称", "门店", "客户名称", "客户"],
    "model": ["型号", "机型", "产品型号", "商品名称", "产品名称"],
    "period": ["月份", "年月", "下单月份", "发放月份", "期间", "账期", "统计月份"],
    "order_date": ["订单日期", "下单日期", "销售日期", "日期"],
    "category": ["品类", "产品品类", "品类名称", "产品类别"],
    "region": ["大区战区", "大区", "战区", "区域"],
    "qty": ["数量", "本次发放数量", "本次下单数量", "下单数量", "发放数量", "零售数量", "销售数量", "净数量"],
    "amount": ["金额", "下单金额", "销售金额", "含税金额"],
    "red_type": ["红包类型", "类型", "红包类别"],
    "unit_amount": ["单台发放金额", "单台金额", "单台红包金额", "单台红包", "红包标准", "政策金额", "标准金额"],
    "total_amount": ["本次发放金额", "发放金额", "红包金额", "红包总额", "合计金额"],
    "cum_qty": ["累计数量", "累计发放数量", "累计红包数量"],
    "policy_start": ["生效起始月", "起始月", "生效月", "开始月份"],
    "policy_end": ["生效截止月", "截止月", "结束月份"],
    "doc_no": ["政策文号", "文号", "政策编号"],
    "base_period": ["基准月", "期初截止月", "截止月", "统计月份"],
    "cum_order": ["累计下单数量", "累计下单", "期初累计下单"],
    "cum_retail": ["前期零售红包数量", "累计零售红包数量", "期初零售红包", "前期零售红包"],
    "cum_order_re": ["累计下单红包数量", "期初下单红包", "累计下单红包"],
    "retail_qty": ["零售红包数量", "零售数量"],
    "order_qty": ["下单红包数量", "下单数量"],
    "retail_unit": ["零售单台金额", "零售红包标准", "零售单台红包"],
    "order_unit": ["下单单台金额", "下单红包标准", "下单单台红包"],
}


def header_key(cell) -> str | None:
    """把表头单元格映射为内部字段名。"""
    if cell is None:
        return None
    s = str(cell).strip().replace(" ", "").replace("\u3000", "")
    if not s:
        return None
    # 精确匹配优先
    for key, aliases in _ALIAS.items():
        if s in aliases:
            return key
    # 其次最长别名优先：单台发放金额 -> unit_amount，不会被「金额」抢走
    best_key, best_len = None, 0
    for key, aliases in _ALIAS.items():
        for a in aliases:
            if a and a in s and len(a) > best_len:
                best_key, best_len = key, len(a)
    return best_key
