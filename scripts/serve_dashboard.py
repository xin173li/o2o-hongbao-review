"""本地看板服务：静态页面 + JSON 接口 + 网页上传导入 + 导出 Excel。

启动：python scripts/serve_dashboard.py --port 8899
"""
from __future__ import annotations

import argparse
import io
import json
import urllib.parse
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from openpyxl import Workbook

from common import ROOT, connect, db_file, ensure_schema, load_settings, log, path_of, use_db

MODULES = [
    {"no": 1, "name": "下单明细", "key": "order"},
    {"no": 2, "name": "红包发放", "key": "issue"},
    {"no": 3, "name": "红包政策", "key": "policy"},
    {"no": 4, "name": "期初累计", "key": "opening"},
]

CHECK_NAMES = {"C1": "政策核销", "C2": "下单红包核销", "C3": "零售红包核销", "C0": "数据检查"}

MODULE_COLUMNS = {
    "order": ["period", "channel_code", "store_code", "store_name", "model_code", "qty", "amount"],
    "issue": ["period", "channel_code", "store_code", "store_name", "model_code", "red_type", "qty",
              "unit_amount", "total_amount"],
    "policy": ["period_start", "model_code", "red_type", "channel_code", "unit_amount", "doc_no"],
    "opening": ["base_period", "channel_code", "store_code", "store_name", "model_code",
                "cum_order_qty", "cum_retail_qty", "cum_order_rqty"],
}
MODULE_HEADERS = {
    "period": "月份", "channel_code": "渠道", "store_code": "门店", "store_name": "门店",
    "model_code": "型号", "red_type": "红包类型", "qty": "数量", "amount": "金额",
    "unit_amount": "单台金额", "total_amount": "红包金额", "period_start": "生效起始月",
    "doc_no": "政策文号", "base_period": "基准月", "cum_order_qty": "累计下单数量",
    "cum_retail_qty": "前期零售红包数量", "cum_order_rqty": "累计下单红包数量",
    "status": "状态", "note": "说明", "diff": "差异", "left_value": "实际值",
    "right_value": "标准值", "period": "月份", "store": "门店", "check_code": "核对项",
}


def q(conn, sql, args=()):
    return [dict(r) for r in conn.execute(sql, args).fetchall()]


def meta(conn):
    cfg = load_settings()["audit"]
    db_rel = str(db_file()).replace(str(ROOT) + "/", "")
    run = conn.execute("SELECT * FROM audit_run ORDER BY run_id DESC LIMIT 1").fetchone()
    periods = [r[0] for r in conn.execute(
        "SELECT DISTINCT period FROM fact_issue ORDER BY period DESC").fetchall()]
    counts = {
        "policy": conn.execute("SELECT COUNT(*) FROM policy").fetchone()[0],
        "order": conn.execute("SELECT COUNT(*) FROM fact_order").fetchone()[0],
        "issue": conn.execute("SELECT COUNT(*) FROM fact_issue").fetchone()[0],
        "opening": conn.execute("SELECT COUNT(*) FROM opening_balance").fetchone()[0],
    }
    return {
        "db_file": db_rel,
        "run_id": run["run_id"] if run else None,
        "periods_span": run["periods"] if run else "",
        "created_at": run["created_at"] if run else "",
        "baseline": cfg["baseline_period"],
        "judge_from": cfg["judge_from_period"],
        "counts": counts,
        "periods": periods,
    }


def summary(conn):
    run = conn.execute("SELECT run_id FROM audit_run ORDER BY run_id DESC LIMIT 1").fetchone()
    if not run:
        return {"run_id": None, "checks": []}
    run_id = run["run_id"]
    out = []
    for code, name in CHECK_NAMES.items():
        rows = conn.execute(
            "SELECT status, COUNT(*) c FROM audit_result "
            "WHERE run_id=? AND check_code=? GROUP BY status", (run_id, code)).fetchall()
        d = {r["status"]: r["c"] for r in rows}
        diff = conn.execute("SELECT COALESCE(SUM(diff),0) FROM audit_result "
                            "WHERE run_id=? AND check_code=? AND status='FAIL'",
                            (run_id, code)).fetchone()[0]
        out.append({"code": code, "name": name, "total": sum(d.values()),
                    "pass": d.get("PASS", 0), "fail": d.get("FAIL", 0),
                    "warn": d.get("WARN", 0), "diff": round(diff, 2)})
    return {"run_id": run_id, "checks": out}


def results(conn, params):
    run = conn.execute("SELECT run_id FROM audit_run ORDER BY run_id DESC LIMIT 1").fetchone()
    if not run:
        return {"total": 0, "rows": [], "left_label": "实际值", "right_label": "标准值"}
    run_id = run["run_id"]

    where = ["run_id = ?"]
    args = [run_id]
    check = params.get("check", "C1")
    if check in CHECK_NAMES:
        where.append("check_code = ?")
        args.append(check)
    for key, col in (("period", "period"), ("channel", "channel_code"),
                     ("store", "store_code"), ("model", "model_code")):
        if params.get(key):
            where.append(f"{col} = ?")
            args.append(params[key])
    if params.get("status") in ("PASS", "FAIL", "WARN"):
        where.append("status = ?")
        args.append(params["status"])
    if params.get("q"):
        where.append("(store_name LIKE ? OR model_code LIKE ? OR note LIKE ?)")
        args += [f"%{params['q']}%"] * 3

    order = "period DESC, CASE status WHEN 'FAIL' THEN 0 WHEN 'WARN' THEN 1 ELSE 2 END, store_name, model_code" \
        if params.get("order", "desc") == "desc" else "period ASC, store_name, model_code"

    total = conn.execute(f"SELECT COUNT(*) FROM audit_result WHERE {' AND '.join(where)}", args).fetchone()[0]
    size = min(int(params.get("size", 50)), 200)
    page = max(int(params.get("page", 1)), 1)
    rows = q(conn, f"""SELECT period, channel_code, store_name, model_code, red_type,
                              left_label, left_value, right_label, right_value, diff, status, note
                       FROM audit_result WHERE {' AND '.join(where)}
                       ORDER BY {order} LIMIT ? OFFSET ?""", args + [size, (page - 1) * size])
    labels = conn.execute(
        "SELECT left_label, right_label FROM audit_result WHERE run_id=? AND check_code=? LIMIT 1",
        (run_id, check if check in CHECK_NAMES else "C1")).fetchone()
    return {"total": total, "page": page, "size": size, "rows": rows,
            "left_label": labels["left_label"] if labels else "实际值",
            "right_label": labels["right_label"] if labels else "标准值"}


def module_rows(conn, params):
    no = int(params.get("no", 1))
    mod = next((m for m in MODULES if m["no"] == no), MODULES[0])
    cols = MODULE_COLUMNS[mod["key"]]
    table = {"order": "fact_order", "issue": "fact_issue",
             "policy": "policy", "opening": "opening_balance"}[mod["key"]]
    period_col = {"order": "period", "issue": "period", "policy": "period_start",
                  "opening": "base_period"}[mod["key"]]

    where, args = [], []
    if params.get("period"):
        where.append(f"t.{period_col} = ?")
        args.append(params["period"])
    if params.get("channel"):
        where.append("t.channel_code = ?")
        args.append(params["channel"])
    if params.get("model"):
        where.append("t.model_code = ?")
        args.append(params["model"])
    if params.get("q"):
        where.append("(t.model_code LIKE ? OR t.store_code LIKE ? OR s.store_name LIKE ?)")
        args += [f"%{params['q']}%"] * 3
    clause = ("WHERE " + " AND ".join(where)) if where else ""

    alias = "t"
    select = ", ".join(f"{alias}.{c}" if c != "store_name" else f"s.store_name" for c in cols)
    join = ""
    if mod["key"] in ("order", "issue", "opening"):
        join = (f" LEFT JOIN dim_store s ON s.store_code = {alias}.store_code "
                f"AND s.channel_code = {alias}.channel_code")
    total = conn.execute(f"SELECT COUNT(*) FROM {table} {alias} {clause}", args).fetchone()[0]
    size = min(int(params.get("size", 50)), 200)
    page = max(int(params.get("page", 1)), 1)
    if params.get("order", "desc") == "desc":
        order = f"{alias}.{period_col} DESC, {alias}.channel_code, {alias}.store_code, {alias}.model_code"
    else:
        order = f"{alias}.{period_col} ASC"
    rows = q(conn, f"SELECT {select} FROM {table} {alias}{join} {clause} ORDER BY {order} LIMIT ? OFFSET ?",
             args + [size, (page - 1) * size])
    return {"total": total, "page": page, "size": size, "rows": rows,
            "columns": [{"key": c, "label": MODULE_HEADERS.get(c, c)} for c in cols],
            "name": mod["name"]}


def dims(conn):
    return {
        "channels": [r[0] for r in conn.execute("SELECT channel_code FROM dim_channel ORDER BY channel_code")],
        "stores": q(conn, "SELECT store_code, store_name, channel_code FROM dim_store ORDER BY store_name"),
        "models": [r[0] for r in conn.execute("SELECT model_code FROM dim_model ORDER BY model_code")],
        "periods": [r[0] for r in conn.execute(
            "SELECT DISTINCT period FROM fact_issue UNION SELECT DISTINCT period FROM fact_order "
            "ORDER BY period DESC")],
    }


def export_xlsx(conn, params):
    wb = Workbook()
    ws = wb.active
    if params.get("type") == "module":
        data = module_rows(conn, {**params, "size": 100000, "page": 1})
        ws.title = data["name"][:30]
        ws.append([c["label"] for c in data["columns"]])
        for row in data["rows"]:
            ws.append([row.get(c["key"]) for c in data["columns"]])
    else:
        data = results(conn, {**params, "size": 100000, "page": 1})
        ws.title = CHECK_NAMES.get(params.get("check", "C1"), "核对结果")
        ws.append(["月份", "渠道", "门店", "型号", "红包类型",
                   data["left_label"], "实际值", data["right_label"], "标准值", "差异", "状态", "说明"])
        for r in data["rows"]:
            ws.append([r["period"], r["channel_code"], r["store_name"], r["model_code"], r["red_type"],
                       r["left_label"], r["left_value"], r["right_label"], r["right_value"],
                       r["diff"], r["status"], r["note"]])
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


class Handler(BaseHTTPRequestHandler):
    server_version = "HongbaoDashboard/1.0"
    web = path_of("web")

    def log_message(self, fmt, *args):  # 请求日志失败不能影响响应
        try:
            super().log_message(fmt, *args)
        except Exception:
            pass

    # ---------- 响应 ----------
    def send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False, default=str).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_bytes(self, data, ctype, filename=None):
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store, max-age=0")
        if filename:
            self.send_header("Content-Disposition",
                             f"attachment; filename*=UTF-8''{urllib.parse.quote(filename)}")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_file(self, path: Path):
        if not path.exists() or path.is_dir():
            index = self.web / "index.html"
            if index.exists():
                return self.send_file(index)
            return self.send_bytes(b"not found", "text/plain", None)
        ctype = {".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8",
                 ".js": "application/javascript; charset=utf-8",
                 ".svg": "image/svg+xml", ".png": "image/png", ".ico": "image/x-icon"}.get(
            path.suffix, "application/octet-stream")
        self.send_bytes(path.read_bytes(), ctype)

    # ---------- 路由 ----------
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path, params = parsed.path, {k: v[0] for k, v in urllib.parse.parse_qs(parsed.query).items()}
        try:
            if path.startswith("/api/"):
                conn = connect()
                ensure_schema(conn)
                if path == "/api/meta":
                    return self.send_json(meta(conn))
                if path == "/api/summary":
                    return self.send_json(summary(conn))
                if path == "/api/results":
                    return self.send_json(results(conn, params))
                if path == "/api/module":
                    return self.send_json(module_rows(conn, params))
                if path == "/api/dims":
                    return self.send_json(dims(conn))
                if path == "/api/modules":
                    return self.send_json({"modules": MODULES})
                if path == "/api/export":
                    data = export_xlsx(conn, params)
                    name = f"{CHECK_NAMES.get(params.get('check','C1'),'核对结果')}_{datetime.now():%Y%m%d%H%M}.xlsx"
                    return self.send_bytes(
                        data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", name)
                return self.send_json({"error": "未知接口"}, 404)
            # 静态：/ 与任意未知中文路径都回首页
            rel = path.lstrip("/") or "index.html"
            target = (self.web / rel).resolve()
            if self.web.resolve() not in target.parents and target != self.web.resolve():
                target = self.web / "index.html"
            return self.send_file(target)
        except Exception as e:  # 兜底，避免连接被直接掐断
            try:
                self.send_json({"error": f"{type(e).__name__}: {e}"}, 500)
            except Exception:
                pass

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        params = {k: v[0] for k, v in urllib.parse.parse_qs(parsed.query).items()}
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length) if length else b""
        try:
            if parsed.path == "/api/recheck":
                from run_audit import main as audit_main
                import sys
                sys.argv = ["run_audit.py"]
                audit_main()
                conn = connect()
                return self.send_json({"ok": True, "summary": summary(conn), "meta": meta(conn)})
            if parsed.path == "/api/import":
                name = params.get("name") or f"上传_{datetime.now():%H%M%S}.xlsx"
                if not name.endswith((".xlsx", ".xlsm")):
                    name += ".xlsx"
                from import_data import import_bytes
                r = import_bytes(body, name)
                out = {"ok": True, "file": r["file"], "written": r["written"],
                       "failed": r["failed"], "batch_id": r["batch_id"],
                       "sheets": r["sheets"]}
                conn = connect()
                out["meta"] = meta(conn)
                out["summary"] = summary(conn)
                return self.send_json(out)
            return self.send_json({"error": "未知接口"}, 404)
        except Exception as e:
            return self.send_json({"error": f"{type(e).__name__}: {e}"}, 500)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8899)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--db", help="指定数据库文件（自己的库）")
    args = parser.parse_args()
    use_db(args.db)
    conn = connect()
    ensure_schema(conn)
    conn.close()
    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    log(f"当前数据库：{db_file()}")
    log(f"看板已启动：http://{args.host}:{args.port}/")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        log("看板已停止")


if __name__ == "__main__":
    main()
