"""执行四项核对（C1 政策 / C2 下单红包 / C3 零售累计 / C0 数据检查）并导出 Excel 报告。"""
from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from common import add_months, connect, db_file, ensure_schema, load_settings, log, path_of, period_range, use_db

CHECK_NAMES = {
    "C1": "政策核销",
    "C2": "下单红包核销",
    "C3": "零售红包核销",
    "C0": "数据检查",
}


def policy_lookup(policies, model, red_type, period, channel):
    cand = [p for p in policies
            if p["model_code"] == model and p["red_type"] == red_type
            and p["period_start"] <= period
            and (p["period_end"] is None or period <= p["period_end"])
            and (p["channel_code"] is None or p["channel_code"] == channel)]
    if not cand:
        return None
    cand.sort(key=lambda p: (p["channel_code"] is not None, p["period_start"]))
    return cand[-1]


def build_rows(conn, cfg):
    baseline = cfg["audit"]["baseline_period"]
    judge_from = cfg["audit"]["judge_from_period"]
    atol = cfg["audit"]["amount_tolerance"]
    qtol = cfg["audit"]["qty_tolerance"]

    policies = [dict(r) for r in conn.execute("SELECT * FROM policy")]
    orders = [dict(r) for r in conn.execute("SELECT * FROM fact_order")]
    issues = [dict(r) for r in conn.execute("SELECT * FROM fact_issue")]
    openings = {(r["channel_code"], r["store_code"], r["model_code"]): dict(r)
                for r in conn.execute("SELECT * FROM opening_balance")}
    store_name = {(r["channel_code"], r["store_code"]): r["store_name"]
                  for r in conn.execute("SELECT * FROM dim_store")}

    rows = []

    def add(check, period, channel, store_code, model, red_type,
            left_label, left_value, right_label, right_value, status, note=""):
        diff = None
        if left_value is not None and right_value is not None:
            diff = round(left_value - right_value, 4)
        rows.append({
            "check_code": check, "period": period, "channel_code": channel,
            "store_code": store_code, "store_name": store_name.get((channel, store_code), store_code),
            "model_code": model or "", "red_type": red_type or "",
            "left_label": left_label, "left_value": left_value,
            "right_label": right_label, "right_value": right_value,
            "diff": diff, "status": status, "note": note,
        })

    # 累计基数：基准月期初 + 基准月之后各月
    def sum_between(items, key, channel, store, model, start, end, extra=None):
        total = 0.0
        for it in items:
            if channel is not None and it["channel_code"] != channel:
                continue
            if it["store_code"] != store:
                continue
            if model is not None and it["model_code"] != model:
                continue
            if key is not None and it.get(key) != extra:
                continue
            if start <= it["period"] <= end:
                total += it["qty"] or 0
        return total

    # ---------- C1 政策核销 ----------
    for it in issues:
        if it["period"] < judge_from:
            continue
        pol = policy_lookup(policies, it["model_code"], it["red_type"], it["period"], it["channel_code"])
        label_type = "零售红包" if it["red_type"] == "retail" else "下单红包"
        left = round(it["unit_amount"] or 0, 4)
        if pol is None:
            add("C1", it["period"], it["channel_code"], it["store_code"], it["model_code"], it["red_type"],
                "实发单台金额（元）", left, "当月政策标准（元）", None, "FAIL",
                f"无当月政策（缺口 {it['period']}，{label_type}）")
            continue
        right = round(pol["unit_amount"], 4)
        if left > right + atol:
            add("C1", it["period"], it["channel_code"], it["store_code"], it["model_code"], it["red_type"],
                "实发单台金额（元）", left, "当月政策标准（元）", right, "FAIL",
                f"高于标准 {round(left - right, 2)} 元")
        else:
            note = "低于标准，按口径不标记异常" if left < right - atol else ""
            add("C1", it["period"], it["channel_code"], it["store_code"], it["model_code"], it["red_type"],
                "实发单台金额（元）", left, "当月政策标准（元）", right, "PASS", note)

    # ---------- C2 下单红包核销 ----------
    order_by_store = {}
    for o in orders:
        order_by_store[(o["period"], o["channel_code"], o["store_code"])] = \
            order_by_store.get((o["period"], o["channel_code"], o["store_code"]), 0) + (o["qty"] or 0)
    issue_order = {}
    for it in issues:
        if it["red_type"] != "order":
            continue
        k = (it["period"], it["channel_code"], it["store_code"])
        issue_order[k] = issue_order.get(k, 0) + (it["qty"] or 0)

    for (period, channel, store), issued in sorted(issue_order.items()):
        if period < judge_from:
            continue
        ordered = order_by_store.get((period, channel, store), 0)
        status = "FAIL" if issued > ordered + qtol else "PASS"
        note = f"超发 {round(issued - ordered, 2)} 台" if status == "FAIL" else ""
        add("C2", period, channel, store, "", "order",
            "当月下单红包发放数量（台）", round(issued, 4),
            "当月净下单数量（台）", round(ordered, 4), status, note)

    # ---------- C3 零售红包核销（累计） ----------
    combos = {(it["period"], it["channel_code"], it["store_code"], it["model_code"])
              for it in issues if it["red_type"] == "retail"}
    combos |= {(o["period"], o["channel_code"], o["store_code"], o["model_code"]) for o in orders}

    for period, channel, store, model in sorted(combos):
        if period < judge_from:
            continue
        op = openings.get((channel, store, model))
        if op:
            base_retail, base_order = op["cum_retail_qty"] or 0, op["cum_order_qty"] or 0
            note_prefix = ""
        else:
            base_retail, base_order = 0, 0
            note_prefix = "期初累计缺失，按 0 起算；"

        cum_retail = base_retail + sum_between(issues, "red_type", channel, store, model,
                                               add_months(baseline, 1), period, extra="retail")
        cum_order = base_order + sum_between(orders, None, channel, store, model,
                                             add_months(baseline, 1), period)

        if cum_retail > cum_order + qtol:
            status, note = "FAIL", note_prefix + f"累计超发 {round(cum_retail - cum_order, 2)} 台"
        elif not op:
            status, note = "WARN", "期初累计缺失，结论待补数据后复核"
        else:
            status, note = "PASS", ""
        add("C3", period, channel, store, model, "retail",
            "累计零售红包发放（台）", round(cum_retail, 4),
            "累计下单数量（台）", round(cum_order, 4), status, note)

    # ---------- C0 数据检查 ----------
    periods_with_issue = {it["period"] for it in issues}
    periods_with_order = {o["period"] for o in orders}
    for p in sorted(periods_with_issue - periods_with_order):
        add("C0", p, "", "", "", "", "该月发放记录（张）",
            len([i for i in issues if i["period"] == p]), "该月下单记录（行）", 0, "WARN",
            "该月缺少下单数据，C2/C3 分母为 0")

    models_in_issue = {it["model_code"] for it in issues}
    models_in_order = {o["model_code"] for o in orders}
    for m in sorted(models_in_issue - models_in_order):
        add("C0", "", "", "", m, "", "有发放的型号（个）", 1, "有下单的型号（个）", 0, "WARN",
            "该型号从未有下单记录，核对③无分母")

    missing_op = sorted({(it["channel_code"], it["store_code"], it["model_code"]) for it in issues
                         if (it["channel_code"], it["store_code"], it["model_code"]) not in openings})
    for channel, store, model in missing_op:
        add("C0", "", channel, store, model, "", "期初累计记录（条）", 0, "应有期初（条）", 1, "WARN",
            "缺期初累计，C3 从 0 起算")

    # C1 无政策也归入数据检查的缺口清单
    gaps = sorted({(it["period"], it["model_code"], it["red_type"]) for it in issues
                   if policy_lookup(policies, it["model_code"], it["red_type"], it["period"],
                                    it["channel_code"]) is None})
    for period, model, red_type in gaps:
        add("C0", period, "", "", model, red_type, "政策缺口（类型）", 1, "应备政策（份）", 0, "WARN",
            f"{period} 无 {model} 的{'零售' if red_type == 'retail' else '下单'}红包政策")

    return rows


def save_run(conn, rows, cfg):
    fingerprint = hashlib.sha256(json.dumps([
        conn.execute("SELECT COUNT(*) FROM policy").fetchone()[0],
        conn.execute("SELECT COUNT(*) FROM fact_order").fetchone()[0],
        conn.execute("SELECT COUNT(*) FROM fact_issue").fetchone()[0],
        conn.execute("SELECT COALESCE(MAX(period),'') FROM fact_issue").fetchone()[0],
        conn.execute("SELECT COALESCE(MAX(period),'') FROM fact_order").fetchone()[0],
    ], sort_keys=True).encode()).hexdigest()[:16]

    periods = sorted({r["period"] for r in rows if r["period"]})
    span = f"{periods[0]} ~ {periods[-1]}" if periods else ""

    # 源数据未变则复用上一批次（覆盖上次结果，不堆重复批次）
    last = conn.execute("SELECT run_id, fingerprint FROM audit_run ORDER BY run_id DESC LIMIT 1").fetchone()
    run_id = None
    if last and last["fingerprint"] == fingerprint:
        run_id = last["run_id"]
        conn.execute("DELETE FROM audit_result WHERE run_id = ?", (run_id,))
        conn.execute("UPDATE audit_run SET created_at = datetime('now'), periods = ? WHERE run_id = ?",
                     (span, run_id))
        log(f"源数据未变，复用批次 #{run_id} 并覆盖上次结果")
    else:
        cur = conn.execute(
            "INSERT INTO audit_run (periods, baseline, judge_from, fingerprint, status) VALUES (?, ?, ?, ?, 'DONE')",
            (span, cfg["audit"]["baseline_period"], cfg["audit"]["judge_from_period"], fingerprint))
        run_id = cur.lastrowid
        log(f"新建核对批次 #{run_id}（{span}）")

    with conn:
        for r in rows:
            conn.execute(
                "INSERT INTO audit_result (run_id, check_code, period, channel_code, store_code, store_name, "
                "model_code, red_type, left_label, left_value, right_label, right_value, diff, status, note) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (run_id, r["check_code"], r["period"], r["channel_code"], r["store_code"], r["store_name"],
                 r["model_code"], r["red_type"], r["left_label"], r["left_value"], r["right_label"],
                 r["right_value"], r["diff"], r["status"], r["note"]))
    return run_id, span


def summarize(conn, run_id):
    out = []
    for code, name in CHECK_NAMES.items():
        q = conn.execute("SELECT status, COUNT(*) FROM audit_result "
                         "WHERE run_id=? AND check_code=? GROUP BY status", (run_id, code)).fetchall()
        d = {r[0]: r[1] for r in q}
        diff = conn.execute("SELECT COALESCE(SUM(diff),0) FROM audit_result "
                            "WHERE run_id=? AND check_code=? AND status='FAIL'", (run_id, code)).fetchone()[0]
        out.append({
            "code": code, "name": name,
            "total": sum(d.values()), "pass": d.get("PASS", 0),
            "fail": d.get("FAIL", 0), "warn": d.get("WARN", 0),
            "diff": round(diff, 2),
        })
    return out


def export_report(conn, run_id, span):
    out = path_of("out")
    path = out / f"红包核对报告_{run_id:03d}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="EEF0F3")
    head_font = Font(bold=True)

    def sheet(title, headers, data, widths=None):
        ws = wb.create_sheet(title)
        ws.append(headers)
        for c in ws[1]:
            c.fill, c.font = head_fill, head_font
            c.alignment = Alignment(vertical="center")
        for row in data:
            ws.append(list(row))
        for i, w in enumerate(widths or [], start=1):
            ws.column_dimensions[chr(64 + i)].width = w
        ws.freeze_panes = "A2"
        return ws

    ws = wb.active
    ws.title = "核对总览"
    ws.append(["O2O 红包核销报告"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([f"批次 #{run_id}", f"核对月份 {span}",
               f"期初基准月 {load_settings()['audit']['baseline_period']}",
               f"生成时间 {datetime.now():%Y-%m-%d %H:%M}"])
    ws.append([])
    ws.append(["核对项", "项数", "合规", "异常", "预警", "差异合计"])
    for c in ws[4]:
        c.fill, c.font = head_fill, head_font
    log(f"当前数据库：{db_file()}")
    for s in summarize(conn, run_id):
        ws.append([s["name"], s["total"], s["pass"], s["fail"], s["warn"], s["diff"]])
    ws.append([])
    ws.append(["数据规模"])
    ws["A" + str(ws.max_row)].font = head_font
    for label, sql in (("政策记录", "SELECT COUNT(*) FROM policy"),
                       ("下单明细", "SELECT COUNT(*) FROM fact_order"),
                       ("发放明细", "SELECT COUNT(*) FROM fact_issue"),
                       ("期初累计", "SELECT COUNT(*) FROM opening_balance")):
        ws.append([label, conn.execute(sql).fetchone()[0]])
    for col, w in (("A", 18), ("B", 14), ("C", 12), ("D", 12), ("E", 12), ("F", 14)):
        ws.column_dimensions[col].width = w

    def fetch(check=None, status=None):
        sql = ("SELECT period, store_name, model_code, red_type, left_label, left_value, right_label, "
               "right_value, diff, status, note FROM audit_result WHERE run_id=?")
        args = [run_id]
        if check:
            sql += " AND check_code=?"
            args.append(check)
        if status:
            sql += " AND status=?"
            args.append(status)
        sql += " ORDER BY period DESC, status, store_name, model_code"
        return conn.execute(sql, args).fetchall()

    headers = ["月份", "门店", "型号", "红包类型", "左侧指标", "实际值", "右侧指标", "标准值", "差异", "状态", "说明"]
    for code, title in (("C1", "核对1 政策核销"), ("C2", "核对2 下单红包"), ("C3", "核对3 零售累计")):
        sheet(title, headers, fetch(code), [10, 24, 12, 10, 20, 10, 18, 10, 10, 8, 32])
    sheet("数据检查", headers, fetch("C0"), [10, 24, 12, 10, 20, 10, 18, 10, 10, 8, 36])

    # 差异汇总：按项×月份×渠道
    agg = conn.execute(
        "SELECT check_code, period, channel_code, COUNT(*), "
        "SUM(CASE WHEN status='FAIL' THEN 1 ELSE 0 END), COALESCE(SUM(diff),0) "
        "FROM audit_result WHERE run_id=? GROUP BY check_code, period, channel_code "
        "ORDER BY check_code, period DESC", (run_id,)).fetchall()
    sheet("差异汇总", ["核对项", "月份", "渠道", "项数", "异常数", "差异合计"],
          [[CHECK_NAMES.get(a[0], a[0]), a[1], a[2] or "—", a[3], a[4], round(a[5], 2)] for a in agg],
          [14, 10, 10, 10, 10, 12])

    # 政策缺口
    gaps = conn.execute(
        "SELECT DISTINCT period, model_code, red_type FROM audit_result "
        "WHERE run_id=? AND check_code='C0' AND left_label LIKE '政策缺口%' ORDER BY period, model_code",
        (run_id,)).fetchall()
    sheet("政策缺口", ["月份", "型号", "红包类型", "处理建议"],
          [[g[0], g[1], "零售红包" if g[2] == "retail" else "下单红包", "补该月政策后重跑核对"]
           for g in gaps], [10, 14, 12, 28])

    # 导入日志
    logs = conn.execute(
        "SELECT batch_id, file_name, sheet_name, data_type, rows_written, rows_failed, errors "
        "FROM import_log ORDER BY batch_id DESC LIMIT 200").fetchall()
    sheet("导入日志", ["批次", "文件", "工作表", "类型", "写入行", "失败行", "错误"],
          [list(l) for l in logs], [8, 30, 18, 12, 10, 10, 60])

    sheet("核对口径说明",
          ["核对项", "判定规则"],
          [["政策核销 C1", f"实发单台高于当月政策标准且超出 {load_settings()['audit']['amount_tolerance']} 元判异常；"
                           "低于或等于标准不标记异常；找不到当月政策判异常"],
           ["下单红包核销 C2", "门店当月下单红包发放数量 > 当月净下单数量（退货为负）判异常"],
           ["零售红包核销 C3", "期初累计 + 其后各月累计零售 > 累计下单判异常；缺期初给预警"],
           ["数据检查 C0", "政策缺口、缺下单数据、型号无下单、缺期初，只预警不判违规"],
           ["状态含义", "PASS 合规 / FAIL 异常 / WARN 预警（数据不完备，需补数复核）"]],
          [20, 90])

    wb.save(path)
    latest = out / "最新核对报告.xlsx"
    wb.save(latest)
    return path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--no-report", action="store_true")
    parser.add_argument("--db", help="指定数据库文件（自己的库）")
    args = parser.parse_args()
    use_db(args.db)

    cfg = load_settings()
    conn = connect()
    ensure_schema(conn)

    if not conn.execute("SELECT COUNT(*) FROM fact_issue").fetchone()[0]:
        log("库里没有发放数据，先跑 python scripts/import_data.py 导入底表", "WARN")
        return

    rows = build_rows(conn, cfg)
    run_id, span = save_run(conn, rows, cfg)
    log(f"当前数据库：{db_file()}")
    for s in summarize(conn, run_id):
        log(f"{s['code']} {s['name']}｜项数 {s['total']}｜合规 {s['pass']}｜异常 {s['fail']}｜预警 {s['warn']}")
    if not args.no_report:
        p = export_report(conn, run_id, span)
        log(f"报告已生成：{p}")
        log(f"同时更新：{path_of('out') / '最新核对报告.xlsx'}")
    conn.close()


if __name__ == "__main__":
    main()
