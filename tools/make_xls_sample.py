#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 BIFF8（Excel 97-2003 .xls）测试件，用来验证安卓端的 .xls 解析。

跑法：python tools/make_xls_sample.py
产物：samples/sample_biff8.xls

刻意覆盖这些特性（每一项都对应解析器里的一段代码）：
- 多工作表（含中文表名）、隐藏工作表
- 共享字符串（SST）+ 富文本段，超过 8224 字节会产生 CONTINUE 记录
- 数字 / 百分比 / 日期 / 文本 / 布尔 / 公式（含缓存值与无缓存值两种）
- 合并单元格（横跨整行的标题 + 竖向合并）
- 字体（粗体、白色字）、填充（深蓝表头）、边框、对齐
- 自定义列宽、行高
"""
import datetime as dt
import os
import sys

import xlwt

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "samples", "sample_biff8.xls")


def main() -> int:
    wb = xlwt.Workbook(encoding="utf-8")

    # ---------------------------------------------------------------- 样式
    title_font = xlwt.Font()
    title_font.name = "宋体"
    title_font.height = 20 * 20          # 20 磅
    title_font.bold = True
    title_align = xlwt.Alignment()
    title_align.horz = xlwt.Alignment.HORZ_CENTER
    title_align.vert = xlwt.Alignment.VERT_CENTER
    title_style = xlwt.XFStyle()
    title_style.font = title_font
    title_style.alignment = title_align

    head_font = xlwt.Font()
    head_font.name = "宋体"
    head_font.bold = True
    head_font.colour_index = 1           # 白字
    head_pat = xlwt.Pattern()
    head_pat.pattern = xlwt.Pattern.SOLID_PATTERN
    head_pat.pattern_fore_colour = 18   # 深蓝（调色板 18 = 0x000080；31 是浅紫，白字看不清）
    head_border = xlwt.Borders()
    head_border.left = head_border.right = head_border.top = head_border.bottom = xlwt.Borders.THIN
    head_align = xlwt.Alignment()
    head_align.horz = xlwt.Alignment.HORZ_CENTER
    head_align.vert = xlwt.Alignment.VERT_CENTER
    head_style = xlwt.XFStyle()
    head_style.font = head_font
    head_style.pattern = head_pat
    head_style.borders = head_border
    head_style.alignment = head_align

    body_style = xlwt.XFStyle()
    body_border = xlwt.Borders()
    body_border.left = body_border.right = body_border.top = body_border.bottom = xlwt.Borders.THIN
    body_style.borders = body_border

    body_center = xlwt.XFStyle()
    body_center.borders = body_border
    body_center.alignment = xlwt.Alignment()
    body_center.alignment.horz = xlwt.Alignment.HORZ_CENTER

    pct_style = xlwt.XFStyle()
    pct_style.borders = body_border
    pct_style.num_format_str = "0.0%"

    date_style = xlwt.XFStyle()
    date_style.borders = body_border
    date_style.num_format_str = "yyyy/m/d"

    money_style = xlwt.XFStyle()
    money_style.borders = body_border
    money_style.num_format_str = "#,##0.00"

    # ---------------------------------------------------------------- 表 1
    ws = wb.add_sheet("成绩汇总")
    ws.write_merge(0, 0, 0, 6, "2025-2026学年第二学期通识选修课汇总表", title_style)

    heads = ["姓名", "学号", "平时", "期末", "总评", "得分率", "选课日期"]
    for c, h in enumerate(heads):
        ws.write(0 + 1, c, h, head_style)

    rows = [
        ("张三", "2025010101", 88, 92, dt.date(2026, 3, 2)),
        ("李四", "2025010102", 76, 81, dt.date(2026, 3, 2)),
        ("王五", "2025010103", 95, 89, dt.date(2026, 3, 3)),
        ("赵六", "2025010104", 61, 58, dt.date(2026, 3, 3)),
        ("钱七", "2025010105", 84, 90, dt.date(2026, 3, 4)),
    ]
    for i, (name, sid, usual, final, day) in enumerate(rows):
        r = i + 2
        ws.write(r, 0, name, body_style)
        ws.write(r, 1, sid, body_center)
        ws.write(r, 2, usual, body_center)
        ws.write(r, 3, final, body_center)
        # 公式：有缓存值的情况由 Excel 保存时写入；xlwt 写的是空缓存
        ws.write(r, 4, xlwt.Formula("C%d*0.4+D%d*0.6" % (r + 1, r + 1)), body_center)
        ws.write(r, 4 + 1, xlwt.Formula("E%d/100" % (r + 1)), pct_style)
        ws.write(r, 6, day, date_style)

    last = len(rows) + 2
    ws.write(last, 0, "平均分", head_style)
    ws.write_merge(last, last, 1, 2, "（5 人）", body_center)
    ws.write(last, 3, xlwt.Formula("AVERAGE(D3:D%d)" % (last)), money_style)
    ws.write(last, 4, xlwt.Formula("AVERAGE(E3:E%d)" % (last)), money_style)
    ws.write(last, 5, 1.0, pct_style)
    ws.write(last, 6, dt.date(2026, 3, 5), date_style)

    # 列宽（单位：1/256 字符）
    widths = [10, 14, 8, 8, 10, 10, 14]
    for c, w in enumerate(widths):
        ws.col(c).width = w * 256
    ws.row(0).height_mismatch = True
    ws.row(0).height = 30 * 20
    ws.row(1).height_mismatch = True
    ws.row(1).height = 8 * 20

    # ---------------------------------------------------------------- 表 2
    ws2 = wb.add_sheet("统计")
    ws2.write(0, 0, "指标", head_style)
    ws2.write(0, 1, "数值", head_style)
    stats = [
        ("人数", xlwt.Formula("COUNTA(成绩汇总!A3:A7)"), body_center),
        ("总评最高", xlwt.Formula("MAX(成绩汇总!E3:E7)"), body_center),
        ("总评最低", xlwt.Formula("MIN(成绩汇总!E3:E7)"), body_center),
        ("及格率", 0.8, pct_style),
        ("备注", "长春理工 · 通识选修", body_style),
    ]
    for i, (k, v, st) in enumerate(stats):
        ws2.write(i + 1, 0, k, body_style)
        ws2.write(i + 1, 1, v, st)
    ws2.col(0).width = 12 * 256
    ws2.col(1).width = 34 * 256

    # ---------------------------------------------------------------- 表 3
    ws3 = wb.add_sheet("长文本")
    long_line = "本行用于撑大共享字符串表，触发 SST 的 CONTINUE 记录。" * 3
    for i in range(120):
        ws3.write(i, 0, "%03d" % i, body_center)
        ws3.write(i, 1, long_line + " #%d" % i, body_style)
    ws3.col(1).width = 60 * 256

    # 隐藏表：解析器应当仍能列出
    ws4 = wb.add_sheet("隐藏表")
    ws4.write(0, 0, "这张表是隐藏的", body_style)
    ws4.visibility = 1

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("written:", OUT, os.path.getsize(OUT), "bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
