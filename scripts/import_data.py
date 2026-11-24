"""Excel 导入器：一个文件里可以有多张表，按表名自动识别类型。

识别规则（可用 config/settings.json 之外的约定，不建议改）：
  表名含「政策」「标准」            -> 红包政策
  表名含「期初」「累计」            -> 期初累计
  表名含「发放」「台账」「红包」    -> 红包发放
  表名含「下单」「订单」「销售」    -> 下单明细

发放表支持两种版式：
  长表：一条一行，有「红包类型」列
  宽表：一行含「零售红包数量」「下单红包数量」两列，自动拆成两行
"""
from __future__ import annotations

import hashlib
import io
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from common import (connect, ensure_schema, header_key, load_settings, log, norm_model, norm_period,
                    norm_text, path_of, red_type_of, to_num)

TYPE_KEYWORDS = [
    ("POLICY", ("政策", "标准")),
    ("OPENING", ("期初", "累计")),
    ("ISSUE", ("发放", "台账", "红包")),
    ("ORDER", ("下单", "订单", "销售")),
]

TYPE_NAME = {"POLICY": "红包政策", "ORDER": "下单明细", "ISSUE": "红包发放", "OPENING": "期初累计"}


def detect_type(sheet_name: str) -> str | None:
    for data_type, words in TYPE_KEYWORDS:
        if any(w in sheet_name for w in words):
            return data_type
    return None


def find_header(rows: list[list]) -> tuple[int, list]:
    """在前 15 行里找表头行：识别字段最多的那一行。"""
    best_idx, best_cols, best_hits = 0, [], 0
    for idx, row in enumerate(rows[:15]):
        cols = [header_key(c) for c in row]
        hits = len([c for c in cols if c])
        if hits > best_hits:
            best_idx, best_cols, best_hits = idx, cols, hits
    return best_idx, best_cols


def norm_channel(value) -> str | None:
    s = norm_text(value)
    if not s:
        return None
    for c in ("京东", "天猫", "零售云"):
        if c in s:
            return c
    if "JD" in s.upper():
        return "京东"
    if "TM" in s.upper():
        return "天猫"
    return s


def _ensure_dim(conn, channel: str, store_code: str, store_name: str, model_raw) -> str:
    model_code = norm_model(model_raw)
    conn.execute(
        "INSERT INTO dim_model (model_code, model_name) VALUES (?, ?) "
        "ON CONFLICT(model_code) DO NOTHING", (model_code, norm_text(model_raw)))
    if store_code:
        conn.execute(
            "INSERT INTO dim_store (store_code, store_name, channel_code) VALUES (?, ?, ?) "
            "ON CONFLICT(store_code, channel_code) DO UPDATE SET store_name = excluded.store_name",
            (store_code, store_name or store_code, channel or "未知"))
    return model_code


def import_sheet(conn, ws, data_type: str, src: str) -> dict:
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    stat = {"sheet": ws.title, "type": data_type, "written": 0, "failed": 0, "errors": []}
    if not rows:
        stat["errors"].append("空工作表")
        return stat

    idx, cols = find_header(rows)
    body = rows[idx + 1:]
    col_pos = {}
    for i, key in enumerate(cols):
        if key and key not in col_pos:
            col_pos[key] = i

    # 政策表单独处理（可能没有「红包类型」列，而是两列标准）
    if data_type == "POLICY":
        return _import_policy(conn, ws, rows, idx, body, cols, src, stat)

    if data_type == "ISSUE" and "red_type" not in col_pos and \
            ("retail_qty" in col_pos or "order_qty" in col_pos):
        return _import_issue_wide(conn, body, cols, col_pos, src, stat)

    required = {
        "ORDER": ("store_name", "model", "period", "qty"),
        "ISSUE": ("store_name", "model", "period", "red_type", "qty"),
        "OPENING": ("store_name", "model", "base_period", "cum_order"),
    }[data_type]
    missing = [k for k in required if k not in col_pos]
    if missing:
        stat["errors"].append(f"缺少必需列：{missing}（识别到的表头：{cols}）")
        return stat

    for r_i, row in enumerate(body, start=idx + 2):
        try:
            if not any(v not in (None, "") for v in row):
                continue
            g = lambda k: row[col_pos[k]] if k in col_pos and col_pos[k] < len(row) else None
            store_name = norm_text(g("store_name"))
            model_raw = g("model")
            if not store_name or not model_raw:
                continue
            channel = norm_channel(g("channel"))
            store_code = norm_text(g("store_code")) or store_name
            period = norm_period(g("period"), default_year=datetime.now().year)

            if data_type == "ORDER":
                if not period:
                    stat["failed"] += 1
                    stat["errors"].append(f"第 {r_i} 行：月份无法识别（{g('period')}）")
                    continue
                qty = to_num(g("qty"))
                if qty is None:
                    stat["failed"] += 1
                    stat["errors"].append(f"第 {r_i} 行：数量为空")
                    continue
                model_code = _ensure_dim(conn, channel, store_code, store_name, model_raw)
                conn.execute(
                    "INSERT INTO fact_order (period, channel_code, store_code, model_code, qty, amount, "
                    "order_date, category, region, src_file) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
                    "ON CONFLICT(period, channel_code, store_code, model_code) DO UPDATE SET "
                    "qty = excluded.qty, amount = excluded.amount, order_date = excluded.order_date, "
                    "category = excluded.category, region = excluded.region, src_file = excluded.src_file",
                    (period, channel or "未知", store_code, model_code, qty, to_num(g("amount")),
                     norm_text(g("order_date")), norm_text(g("category")), norm_text(g("region")), src))
                stat["written"] += 1

            elif data_type == "ISSUE":
                if not period:
                    stat["failed"] += 1
                    stat["errors"].append(f"第 {r_i} 行：月份无法识别")
                    continue
                types = red_type_of(g("red_type"))
                if not types:
                    stat["failed"] += 1
                    stat["errors"].append(f"第 {r_i} 行：红包类型无法识别（{g('red_type')}）")
                    continue
                qty = to_num(g("qty")) or 0
                unit = to_num(g("unit_amount"))
                total = to_num(g("total_amount"))
                if unit is None and qty:
                    unit = (total or 0) / qty if total is not None else None
                if unit is None:
                    stat["failed"] += 1
                    stat["errors"].append(f"第 {r_i} 行：单台金额与红包金额都为空")
                    continue
                model_code = _ensure_dim(conn, channel, store_code, store_name, model_raw)
                for rt in types:   # 混写类型：两类分别计入
                    conn.execute(
                        "INSERT INTO fact_issue (period, channel_code, store_code, model_code, red_type, "
                        "qty, unit_amount, total_amount, cum_qty_src, category, region, src_file) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
                        "ON CONFLICT(period, channel_code, store_code, model_code, red_type) DO UPDATE SET "
                        "qty = excluded.qty, unit_amount = excluded.unit_amount, "
                        "total_amount = excluded.total_amount, cum_qty_src = excluded.cum_qty_src, "
                        "category = excluded.category, region = excluded.region",
                        (period, channel or "未知", store_code, model_code, rt, qty, unit,
                         total if total is not None else qty * unit, to_num(g("cum_qty")),
                         norm_text(g("category")), norm_text(g("region")), src))
                    stat["written"] += 1

            elif data_type == "OPENING":
                base = norm_period(g("base_period"), default_year=datetime.now().year)
                if not base:
                    stat["failed"] += 1
                    stat["errors"].append(f"第 {r_i} 行：基准月无法识别")
                    continue
                model_code = _ensure_dim(conn, channel, store_code, store_name, model_raw)
                conn.execute(
                    "INSERT INTO opening_balance (base_period, store_code, model_code, channel_code, "
                    "cum_order_qty, cum_retail_qty, cum_order_rqty, src_file) VALUES (?, ?, ?, ?, ?, ?, ?, ?) "
                    "ON CONFLICT(base_period, channel_code, store_code, model_code) DO UPDATE SET "
                    "cum_order_qty = excluded.cum_order_qty, cum_retail_qty = excluded.cum_retail_qty, "
                    "cum_order_rqty = excluded.cum_order_rqty",
                    (base, store_code, model_code, channel or "未知",
                     to_num(g("cum_order")) or 0, to_num(g("cum_retail")) or 0,
                     to_num(g("cum_order_re")) or 0, src))
                stat["written"] += 1

        except Exception as e:  # 单行失败不影响整表
            stat["failed"] += 1
            stat["errors"].append(f"第 {r_i} 行：{e}")

    return stat


def _import_policy(conn, ws, rows, idx, body, cols, src, stat) -> dict:
    """政策表：优先识别「零售/下单」两类标准列；否则按长表（型号+类型+金额）。"""
    header = [str(c or "") for c in rows[idx]]
    retail_col = next((i for i, h in enumerate(header)
                       if "零售" in h and ("红包" in h or "标准" in h or "金额" in h)), None)
    order_col = next((i for i, h in enumerate(header)
                      if ("下单" in h or "预售" in h) and ("红包" in h or "标准" in h or "金额" in h)), None)
    model_col = next((i for i, h in enumerate(header) if "型号" in h or "机型" in h), None)
    doc_col = next((i for i, h in enumerate(header) if "文号" in h), None)
    channel_col = next((i for i, h in enumerate(header) if "渠道" in h), None)
    type_col = next((i for i, h in enumerate(header) if "红包类型" in h or h.strip() == "类型"), None)
    unit_col = next((i for i, h in enumerate(header) if "单台" in h or "标准金额" in h), None)

    if model_col is None:
        stat["errors"].append("政策表缺少「型号」列")
        return stat

    # 期间：优先逐行取「生效起始月」列，其次取文件名里的「N月」
    period_col = next((i for i, h in enumerate(header)
                       if "生效" in h or "月份" in h or "年月" in h), None)
    file_period = None
    for p in range(1, 13):
        if f"{p}月" in ws.title or f"{p}月" in src:
            file_period = f"{datetime.now().year:04d}-{p:02d}"
            break

    wrote_any = False
    for r_i, row in enumerate(body, start=idx + 2):
        if not any(v not in (None, "") for v in row):
            continue
        model_raw = row[model_col] if model_col < len(row) else None
        if not model_raw:
            continue
        period = None
        if period_col is not None and period_col < len(row):
            period = norm_period(row[period_col], default_year=datetime.now().year)
        period = period or file_period
        if not period:
            stat["failed"] += 1
            stat["errors"].append(f"第 {r_i} 行：政策月份无法确定（无「生效起始月」列且文件名未含「N月」）")
            continue
        model_code = norm_model(model_raw)
        _ensure_dim(conn, None, "", "", model_raw)
        doc = row[doc_col] if doc_col is not None and doc_col < len(row) else None
        channel = norm_channel(row[channel_col]) if channel_col is not None and channel_col < len(row) else None

        if type_col is not None and unit_col is not None:
            # 长表
            types = red_type_of(row[type_col] if type_col < len(row) else None)
            amount = to_num(row[unit_col] if unit_col < len(row) else None)
            items = [(t, amount) for t in types]
        else:
            items = []
            if retail_col is not None:
                items.append(("retail", to_num(row[retail_col] if retail_col < len(row) else None)))
            if order_col is not None:
                items.append(("order", to_num(row[order_col] if order_col < len(row) else None)))

        for red_type, amount in items:
            if amount is None:
                stat["failed"] += 1
                stat["errors"].append(f"第 {r_i} 行：{model_code} 的 {red_type} 标准为空")
                continue
            validity = load_settings()["source_rules"].get("policy_validity", "month")
            period_end = period if validity == "month" else None
            # SQLite 的唯一约束不覆盖 NULL，政策表渠道为空时先删旧行再写，避免重复政策
            conn.execute(
                "DELETE FROM policy WHERE model_code = ? AND red_type = ? "
                "AND IFNULL(channel_code, '*') = ? AND period_start = ?",
                (model_code, red_type, channel or "*", period))
            conn.execute(
                "INSERT INTO policy (model_code, red_type, channel_code, period_start, period_end, "
                "unit_amount, doc_no, src_file) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (model_code, red_type, channel, period, period_end, amount, norm_text(doc), src))
            stat["written"] += 1
            wrote_any = True

    if not wrote_any:
        stat["errors"].append("政策表未写入任何记录，请检查表头是否为「型号 / 总部预售红包｜大区下单红包 / 大区零售红包」")
    return stat


def _import_issue_wide(conn, body, cols, col_pos, src, stat) -> dict:
    """宽表发放：零售与下单各一列数量（可各配一列单台金额）。"""
    need = ("store_name", "model", "period")
    missing = [k for k in need if k not in col_pos]
    if missing:
        stat["errors"].append(f"宽表发放缺少列：{missing}")
        return stat

    for r_i, row in enumerate(body, start=2):
        if not any(v not in (None, "") for v in row):
            continue
        g = lambda k: row[col_pos[k]] if k in col_pos and col_pos[k] < len(row) else None
        store_name, model_raw = norm_text(g("store_name")), g("model")
        period = norm_period(g("period"), default_year=datetime.now().year)
        if not store_name or not model_raw or not period:
            stat["failed"] += 1
            stat["errors"].append(f"第 {r_i} 行：门店/型号/月份不完整")
            continue
        channel = norm_channel(g("channel"))
        store_code = norm_text(g("store_code")) or store_name
        model_code = _ensure_dim(conn, channel, store_code, store_name, model_raw)

        for red_type, qty_key, unit_key in (("retail", "retail_qty", "retail_unit"),
                                            ("order", "order_qty", "order_unit")):
            qty = to_num(g(qty_key))
            if not qty:
                continue
            unit = to_num(g(unit_key)) or 0
            conn.execute(
                "INSERT INTO fact_issue (period, channel_code, store_code, model_code, red_type, "
                "qty, unit_amount, total_amount, src_file) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?) "
                "ON CONFLICT(period, channel_code, store_code, model_code, red_type) DO UPDATE SET "
                "qty = excluded.qty, unit_amount = excluded.unit_amount, total_amount = excluded.total_amount",
                (period, channel or "未知", store_code, model_code, red_type, qty, unit,
                 qty * unit, src))
            stat["written"] += 1
    return stat


def import_file(path: Path, conn=None) -> dict:
    own = conn is None
    if own:
        conn = connect()
        ensure_schema(conn)

    data = path.read_bytes()
    file_hash = hashlib.sha256(data).hexdigest()[:16]
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=False)

    with conn:
        cur = conn.execute(
            "INSERT INTO import_batch (file_name, file_hash, sheet_count, status, summary) "
            "VALUES (?, ?, ?, 'RUNNING', '')", (path.name, file_hash, len(wb.sheetnames)))
        batch_id = cur.lastrowid

    stats, written, failed, sheets_used = [], 0, 0, 0
    for ws in wb.worksheets:
        data_type = detect_type(ws.title)
        if data_type is None:
            continue
        sheets_used += 1
        with conn:
            st = import_sheet(conn, ws, data_type, path.name)
        stats.append(st)
        written += st["written"]
        failed += st["failed"]

    with conn:
        conn.execute(
            "UPDATE import_batch SET finished_at = datetime('now'), rows_written = ?, rows_failed = ?, "
            "status = ?, summary = ? WHERE batch_id = ?",
            (written, failed, "DONE" if written else "EMPTY",
             f"工作表 {sheets_used} 个；写入 {written} 行", batch_id))
        for st in stats:
            conn.execute(
                "INSERT INTO import_log (batch_id, file_name, sheet_name, data_type, rows_written, "
                "rows_failed, errors) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (batch_id, path.name, st["sheet"], TYPE_NAME.get(st["type"], st["type"]),
                 st["written"], st["failed"], "\n".join(st["errors"][:50])))

    if own:
        conn.close()
    return {"file": path.name, "batch_id": batch_id, "written": written,
            "failed": failed, "sheets": stats, "skipped_sheets": len(wb.sheetnames) - sheets_used}


def import_dir(dir_path: Path | None = None) -> list[dict]:
    inbox = dir_path or path_of("inbox")
    files = sorted([p for p in inbox.glob("*.xlsx") if not p.name.startswith("~$")])
    if not files:
        log(f"{inbox} 下没有待导入的 xlsx")
        return []
    conn = connect()
    ensure_schema(conn)
    results = []
    for f in files:
        log(f"导入 {f.name} …")
        r = import_file(f, conn)
        results.append(r)
        log(f"  → 写入 {r['written']} 行，失败 {r['failed']} 行，批次 #{r['batch_id']}")
        for st in r["sheets"]:
            if st["errors"]:
                for e in st["errors"][:5]:
                    log(f"    {st['sheet']}: {e}", "WARN")
        archive = path_of("archive")
        shutil.move(str(f), str(archive / f"{datetime.now():%Y%m%d%H%M%S}_{f.name}"))
    conn.close()
    return results


def import_bytes(data: bytes, filename: str) -> dict:
    """供网页上传使用：先落盘到 inbox，再走同一套导入逻辑。"""
    inbox = path_of("inbox")
    target = inbox / filename
    target.write_bytes(data)
    conn = connect()
    ensure_schema(conn)
    r = import_file(target, conn)
    conn.close()
    return r


if __name__ == "__main__":
    import argparse
    from common import use_db
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", help="指定数据库文件（自己的库）")
    parser.add_argument("--dir", help="指定待导入目录，默认 data/inbox")
    a = parser.parse_args()
    use_db(a.db)
    from pathlib import Path as _Path
    import_dir(_Path(a.dir) if a.dir else None)
