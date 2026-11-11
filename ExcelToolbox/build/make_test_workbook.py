#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成「Excel 小工具箱」测试工作簿（跨平台，只依赖 openpyxl）。

用法： python build/make_test_workbook.py [输出路径]
默认输出： tests/ExcelToolbox_TestData.xlsx

数据经过刻意设计，便于自动化断言：
  * 删除空行测试：共 10 行，其中第 4、7、8 行完全为空  -> 期望删掉 3 行
  * 去重测试    ：共 8 行数据（无表头），4 种组合      -> 期望保留 4 行
  * 美化表头测试：A1:E1 为表头                        -> 期望蓝底白字加粗
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_OUT = ROOT / "tests" / "ExcelToolbox_TestData.xlsx"


def build(out_path: Path) -> None:
    wb = Workbook()

    # ------------------------------------------------------------------
    # 1) 删除空行：第 4、7、8 行完全留空
    # ------------------------------------------------------------------
    ws = wb.active
    ws.title = "删除空行测试"
    rows = [
        ["编号", "姓名", "部门", "金额"],
        [1, "张三", "技术部", 8000],
        [2, "李四", "市场部", 9000],
        None,                                   # 第 4 行：空
        [3, "王五", "财务部", 7000],
        [4, "赵六", "人事部", 6500],
        None,                                   # 第 7 行：空
        None,                                   # 第 8 行：空
        [5, "孙七", "技术部", 8800],
        [6, "周八", "市场部", 9200],
    ]
    for r, values in enumerate(rows, start=1):
        if values is None:
            continue
        for c, value in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=value)
    # 给表头一点原始格式，方便对比「美化表头」的效果
    for c in range(1, 5):
        ws.cell(row=1, column=c).font = Font(bold=False, color="FF000000")
    for col, width in zip("ABCD", (10, 12, 14, 12)):
        ws.column_dimensions[col].width = width

    # ------------------------------------------------------------------
    # 2) 去重：无表头，8 行中 4 行为重复（故意不加序号列，否则每行都唯一）
    # ------------------------------------------------------------------
    ws2 = wb.create_sheet("去重测试")
    dedup = [
        ["张三", "北京", 100],
        ["李四", "上海", 200],
        ["张三", "北京", 100],   # dup of row 1
        ["王五", "广州", 300],
        ["李四", "上海", 200],   # dup of row 2
        ["赵六", "深圳", 400],
        ["张三", "北京", 100],   # dup of row 1
        ["王五", "广州", 300],   # dup of row 4
    ]
    for r, values in enumerate(dedup, start=1):
        for c, value in enumerate(values, start=1):
            ws2.cell(row=r, column=c, value=value)
    for col, width in zip("ABC", (12, 12, 10)):
        ws2.column_dimensions[col].width = width

    # ------------------------------------------------------------------
    # 3) 美化表头
    # ------------------------------------------------------------------
    ws3 = wb.create_sheet("美化表头测试")
    header = ["员工编号", "姓名", "所属部门", "入职日期", "备注说明"]
    for c, value in enumerate(header, start=1):
        ws3.cell(row=1, column=c, value=value)
    data = [
        ["E1001", "张三", "技术研发中心", "2021-03-01", "正式员工"],
        ["E1002", "李四", "市场推广部", "2022-07-15", "正式员工"],
        ["E1003", "王五", "财务管理部", "2020-11-02", "部门负责人"],
        ["E1004", "赵六", "人力资源部", "2023-01-09", "试用期"],
        ["E1005", "孙七", "技术研发中心", "2019-05-20", "资深工程师"],
    ]
    for r, values in enumerate(data, start=2):
        for c, value in enumerate(values, start=1):
            ws3.cell(row=r, column=c, value=value)
    for c in range(1, 6):
        ws3.cell(row=1, column=c).fill = PatternFill(fill_type=None)

    # ------------------------------------------------------------------
    # 4) 说明
    # ------------------------------------------------------------------
    ws4 = wb.create_sheet("说明")
    notes = [
        "Excel 小工具箱 - 测试用工作簿",
        "",
        "【删除空行测试】选中该工作表后点「删除空行」",
        "  共 10 行，第 4 / 7 / 8 行完全为空，应删除 3 行，剩余 7 行。",
        "  执行后应新增一张橙色标签的「备份_删除空行_xxx」工作表。",
        "",
        "【去重测试】选中 A1:C8 后点「一键去重」，首行是否表头请选择「否」",
        "  8 行中有 4 条重复，应保留 4 行：张三/北京/100、李四/上海/200、",
        "  王五/广州/300、赵六/深圳/400（顺序不变）。",
        "",
        "【美化表头测试】选中 A1:E6 后点「美化表头」",
        "  首行应变为蓝色背景 + 白色粗体 + 居中，列宽自动调整。",
        "",
        "注意：VBA 操作不进入 Excel 撤销栈，请用自动生成的备份工作表恢复。",
    ]
    for r, text in enumerate(notes, start=1):
        ws4.cell(row=r, column=1, value=text)
    ws4.column_dimensions["A"].width = 78
    ws4["A1"].font = Font(bold=True, size=13)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def self_check(path: Path) -> None:
    wb = load_workbook(path, data_only=False)
    assert wb.sheetnames == ["删除空行测试", "去重测试", "美化表头测试", "说明"], wb.sheetnames

    ws = wb["删除空行测试"]
    assert ws.max_row == 10 and ws.max_column == 4, (ws.max_row, ws.max_column)
    blanks = [r for r in range(1, 11) if all(ws.cell(r, c).value is None for c in range(1, 5))]
    assert blanks == [4, 7, 8], blanks

    ws2 = wb["去重测试"]
    assert ws2.max_row == 8 and ws2.max_column == 3
    keys = [tuple(ws2.cell(r, c).value for c in range(1, 4)) for r in range(1, 9)]
    assert len(set(keys)) == 4, keys

    ws3 = wb["美化表头测试"]
    assert [ws3.cell(1, c).value for c in range(1, 6)] == [
        "员工编号", "姓名", "所属部门", "入职日期", "备注说明"
    ]

    print(f"[OK] 测试工作簿生成并自检通过: {path}")
    print(f"     工作表: {wb.sheetnames}")
    print(f"     删除空行测试: 10 行 / 空行 {blanks} / 期望删除 3 行")
    print(f"     去重测试    : 8 行 / 唯一组合 {len(set(keys))} 种 / 期望保留 4 行")
    print(f"     美化表头测试: 表头 {[ws3.cell(1, c).value for c in range(1, 6)]}")


if __name__ == "__main__":
    target = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else DEFAULT_OUT
    build(target)
    self_check(target)
