"""生成一份可直接导入的样例底表（含刻意构造的三类异常），用于演示与自检。

产出：samples/样例底表_2026H1.xlsx
  工作表：下单明细 / 红包发放台账 / 红包政策 / 期初累计

注意：本文件内所有门店、型号、政策文号、金额均为虚构示例，不含任何真实业务数据。
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from common import ROOT, period_range

# 虚构型号
MODELS = ["65X1A", "75X1A", "85X2B", "98X2B"]
# 虚构渠道与门店
STORES = [
    ("渠道A", "A001", "示例门店A"),
    ("渠道A", "A002", "示例门店B"),
    ("渠道B", "B001", "示例门店C"),
]
PERIODS = period_range("2026-01", "2026-06")

# 每型号每月下单基数；个别门店/月份故意做小，用来触发 C2 / C3 异常
ORDER_BASE = {"65X1A": 10, "75X1A": 8, "85X2B": 6, "98X2B": 4}
# 虚构红包标准（下单红包, 零售红包）
POLICY = {"65X1A": (15, 30), "75X1A": (15, 30), "85X2B": (18, 36), "98X2B": (22, 44)}


def net_order(code, m, p):
    """净下单数量：退货为负，个别组合刻意做小。"""
    qty = ORDER_BASE[m]
    if code == "A002" and m == "75X1A" and p == "2026-04":
        qty = 8          # 下单 8，下单红包发 12 -> C2 异常
    if code == "A001" and m == "85X2B":
        qty = 4          # 下单小、零售大 -> C3 累计超发
    if code == "B001" and p == "2026-02":
        qty -= 1         # 退货，净数量
    return qty


def write_order(ws):
    ws.append(["渠道", "门店编码", "门店名称", "型号", "下单月份", "下单数量", "下单金额"])
    for ch, code, name in STORES:
        for p in PERIODS:
            for m in MODELS:
                qty = net_order(code, m, p)
                ws.append([ch, code, name, m, p, qty, qty * 1000])


def write_policy(ws):
    ws.append(["型号", "生效起始月", "总部预售红包/大区下单红包", "大区零售红包", "政策文号"])
    for p in PERIODS:
        for m in MODELS:
            if m == "98X2B" and p == "2026-05":
                continue     # 故意缺一个月政策 -> C1「无当月政策」异常
            order_amt, retail_amt = POLICY[m]
            ws.append([m, p, order_amt, retail_amt, f"DEMO-{p}-{m}"])


def write_issue(ws):
    ws.append(["发放月份", "渠道", "门店编码", "门店名称", "型号", "红包类型",
               "发放数量", "单台金额", "红包金额"])
    for ch, code, name in STORES:
        for p in PERIODS:
            for m in MODELS:
                order_amt, retail_amt = POLICY[m]
                oq = net_order(code, m, p)
                rq = max(ORDER_BASE[m] - 2, 1)
                retail_unit = retail_amt

                if code == "A001" and m == "75X1A" and p == "2026-03":
                    retail_unit = retail_amt + 2    # 高于标准 2 元 -> C1 异常
                if code == "A002" and m == "75X1A" and p == "2026-04":
                    oq = 12                        # 下单红包 12 > 当月下单 8 -> C2 异常
                if code == "A001" and m == "85X2B":
                    rq = 8                         # 零售 8 > 下单 4 -> C3 累计超发
                if code == "B001" and m == "98X2B" and p == "2026-05":
                    ws.append([p, ch, code, name, m, "下单红包+零售红包",  # 混写类型
                               3, order_amt, None])
                    continue

                ws.append([p, ch, code, name, m, "下单红包", oq, order_amt, oq * order_amt])
                ws.append([p, ch, code, name, m, "零售红包", rq, retail_unit, rq * retail_unit])


def write_opening(ws):
    ws.append(["基准月", "渠道", "门店编码", "门店名称", "型号",
               "累计下单数量", "前期零售红包数量", "累计下单红包数量"])
    for ch, code, name in STORES:
        for m in MODELS:
            ws.append(["2025-12", ch, code, name, m, 30, 24, 18])


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "下单明细"
    write_order(ws)
    write_policy(wb.create_sheet("红包政策"))
    write_issue(wb.create_sheet("红包发放台账"))
    write_opening(wb.create_sheet("期初累计"))

    out = ROOT / "samples" / "样例底表_2026H1.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"样例底表已生成：{out}")


if __name__ == "__main__":
    main()
