"""生成四张独立底表模板（templates/），每张含「填写说明」页。

产出：
  templates/1_下单明细_模板.xlsx
  templates/2_红包发放_模板.xlsx
  templates/3_红包政策_模板.xlsx
  templates/4_期初累计_模板.xlsx

注意：本文件内的示例门店、型号、品类、政策文号、金额均为虚构示范值，不含真实业务数据。
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from common import ROOT

HEAD_FILL = PatternFill("solid", fgColor="EEF0F3")
HEAD_FONT = Font(bold=True)

SPECS = [
    {
        "file": "1_下单明细_模板.xlsx",
        "sheet": "下单明细",
        "headers": ["月份", "订单日期", "渠道", "门店名称", "型号", "品类", "数量", "金额", "类型", "大区战区"],
        "rows": [
            ["2026-07", "2026-07-22", "渠道A", "示例门店甲", "65X1A", "显示类", 1, 0.00, "下单", "示例大区"],
            ["2026-07", "2026-07-22", "渠道A", "示例门店甲", "X2 Pro", "智能家居", 1, 0.00, "下单", "示例大区"],
        ],
        "widths": [12, 12, 10, 34, 18, 10, 10, 12, 10, 12],
        "notes": [
            ("月份", "是", "2026-07 / 2026年7月 / 202607 均可；退货行仍填原月份", "2026-07"),
            ("订单日期", "否", "现有系统导出的日期列，原样保留即可，不参与核对", "2026-07-22"),
            ("渠道", "是", "渠道A / 渠道B / 渠道C；导出里带其他写法也能识别，新渠道自动建档", "渠道A"),
            ("门店名称", "是", "导出里的客户/门店名，与发放、期初表保持同一写法", "示例门店甲"),
            ("型号", "是", "支持“65X1A”“X2 Pro”“物联网智能锁X3 Pro”等带空格写法，统一按去空格大写归一", "65X1A"),
            ("品类", "否", "显示类 / 智能家居 / 大家电等，原样保留，用于筛选与追溯", "显示类"),
            ("数量", "是", "净数量：退货写负数，不要单开退货数量列", "1"),
            ("金额", "否", "导出里的下单金额，0.00 也照原样填，不影响核对", "0.00"),
            ("类型", "否", "导出里的“下单/退货”等业务类型，原样保留", "下单"),
            ("大区战区", "否", "导出里的组织维度，原样保留，便于以后按大区筛选", "示例大区"),
        ],
    },
    {
        "file": "2_红包发放_模板.xlsx",
        "sheet": "红包发放台账",
        "headers": ["月份", "渠道", "门店名称", "型号", "品类", "红包类型",
                    "本次发放数量", "单台发放金额", "本次发放金额"],
        "rows": [
            ["2026-09", "渠道A", "示例门店乙", "85X2B", "显示类", "零售红包", 1, 30.00, 30.00],
            ["2026-09", "渠道A", "示例门店乙", "85X2B", "显示类", "下单红包", 1, 15.00, 15.00],
        ],
        "widths": [12, 10, 34, 18, 10, 12, 14, 14, 14],
        "notes": [
            ("月份", "是", "红包归属月；跨月台账请拆成每月一行", "2026-09"),
            ("渠道", "是", "渠道A / 渠道B / 渠道C", "渠道A"),
            ("门店名称", "是", "必须能与下单表对上，否则核对的分子会落到空门店", "示例门店乙"),
            ("型号", "是", "与下单表同一写法即可，系统按去空格大写归一", "85X2B"),
            ("品类", "否", "显示类 / 智能家居等，原样保留", "显示类"),
            ("红包类型", "是", "填“零售红包”或“下单红包”；一行含两类时按两类分别计入", "零售红包"),
            ("本次发放数量", "是", "净数量：退货写负数", "1"),
            ("单台发放金额", "是", "实发单台金额（元），核销①就是拿它比政策标准", "30.00"),
            ("本次发放金额", "否", "本次发放总额；不填按“数量×单台金额”推算", "30.00"),
        ],
    },
    {
        "file": "3_红包政策_模板.xlsx",
        "sheet": "红包政策",
        "headers": ["型号", "生效起始月", "总部预售红包/大区下单红包", "大区零售红包", "政策文号"],
        "rows": [
            ["65X1A", "2026-07", 15, 30, "DEMO-2026-07-65X1A"],
            ["85X2B", "2026-09", 18, 36, "DEMO-2026-09-85X2B"],
        ],
        "widths": [18, 14, 26, 14, 34],
        "notes": [
            ("型号", "是", "一个型号一个月一行；不同型号政策不同就各写一行", "65X1A"),
            ("生效起始月", "是", "政策只对当月生效，跨月需再给一行；某月缺失会直接判异常并列出缺口", "2026-07"),
            ("总部预售红包/大区下单红包", "是", "下单红包单台标准（元）", "15"),
            ("大区零售红包", "是", "零售红包单台标准（元）", "30"),
            ("政策文号", "否", "写政策文件编号或标题，便于追溯", "DEMO-2026-07-65X1A"),
        ],
    },
    {
        "file": "4_期初累计_模板.xlsx",
        "sheet": "期初累计",
        "headers": ["基准月", "渠道", "门店名称", "型号", "品类",
                    "累计下单数量", "前期零售红包数量", "累计下单红包数量"],
        "rows": [
            ["2025-12", "渠道A", "示例门店甲", "65X1A", "显示类", 30, 24, 18],
            ["2025-12", "渠道A", "示例门店乙", "85X2B", "显示类", 12, 10, 6],
        ],
        "widths": [12, 10, 34, 18, 10, 16, 20, 20],
        "notes": [
            ("基准月", "是", "期初数据截止到哪个月底，通常是核对起始月的上一个月", "2025-12"),
            ("门店名称", "是", "与下单、发放表同一写法", "示例门店甲"),
            ("型号", "是", "一个门店一个型号一行", "65X1A"),
            ("品类", "否", "原样保留，便于追溯", "显示类"),
            ("累计下单数量", "是", "基准月为止的累计净下单数量，核销③的分母起点", "30"),
            ("前期零售红包数量", "是", "基准月为止的累计零售红包发放数量，核销③的分子起点", "24"),
            ("累计下单红包数量", "否", "基准月为止的累计下单红包发放数量，用于交叉校验", "18"),
        ],
    },
]


def build_combined() -> None:
    """把四张表合成一个 Excel（samples/底表模板_空.xlsx），每张表只留表头 + 一行示例。"""
    wb = Workbook()
    wb.remove(wb.active)
    for spec in SPECS:
        ws = wb.create_sheet(spec["sheet"])
        ws.append(spec["headers"])
        for c in ws[1]:
            c.font, c.fill = HEAD_FONT, HEAD_FILL
        if spec["rows"]:
            ws.append(spec["rows"][0])
        for i, w in enumerate(spec.get("widths") or [], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    out = ROOT / "samples" / "底表模板_空.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"已生成合表模板：{out.name}")


def build() -> None:
    out_dir = ROOT / "templates"
    out_dir.mkdir(parents=True, exist_ok=True)

    for spec in SPECS:
        wb = Workbook()
        ws = wb.active
        ws.title = spec["sheet"]
        ws.append(spec["headers"])
        for c in ws[1]:
            c.font, c.fill = HEAD_FONT, HEAD_FILL
            c.alignment = Alignment(horizontal="center", vertical="center")
        for row in spec["rows"]:
            ws.append(row)
        for i, w in enumerate(spec["widths"], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

        note = wb.create_sheet("填写说明")
        note.append(["字段", "是否必填", "填写要求", "示例"])
        for c in note[1]:
            c.font, c.fill = HEAD_FONT, HEAD_FILL
        for item in spec["notes"]:
            note.append(list(item))
        for i, w in enumerate([22, 10, 56, 26], start=1):
            note.column_dimensions[get_column_letter(i)].width = w
        note.freeze_panes = "A2"

        path = out_dir / spec["file"]
        wb.save(path)
        print(f"已生成模板：{path.name}")


if __name__ == "__main__":
    build()
    build_combined()
