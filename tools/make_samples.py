"""生成演示 / 自测用的样例文件：PDF、Word、Excel、PPT 各一个。"""
from __future__ import annotations

from pathlib import Path

import pymupdf
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = Path(__file__).resolve().parents[1] / "samples"


def make_pdf(path: Path, pages: int = 6) -> None:
    doc = pymupdf.open()
    for i in range(1, pages + 1):
        page = doc.new_page(width=595, height=842)
        page.draw_rect(pymupdf.Rect(40, 40, 555, 96), color=(0.18, 0.44, 0.93), fill=(0.93, 0.96, 1.0))
        page.insert_text((58, 74), f"PagePull 样例报告  ·  第 {i} 页", fontsize=18, fontname="china-s",
                         color=(0.1, 0.15, 0.25))
        page.insert_text((58, 128), f"章节 {i}：季度经营情况摘要", fontsize=13, fontname="china-s")
        body = [
            "本页为演示数据，用于验证按页提取、渲染清晰度与导出格式。",
            "· 支持非连续页码选择，例如只取第 1、3、5 页；",
            "· 导出为 PNG / JPG 时按 DPI 重采样，300 DPI 可用于印刷；",
            "· 导出为 PDF 时保留原始矢量内容与链接。",
        ]
        y = 160
        for line in body:
            page.insert_text((58, y), line, fontsize=11, fontname="china-s", color=(0.25, 0.29, 0.35))
            y += 24

        # 画一个 4×3 的示意表格
        x0, y0, w, h = 58, 300, 480, 120
        page.draw_rect(pymupdf.Rect(x0, y0, x0 + w, y0 + h), color=(0.8, 0.84, 0.9))
        for r in range(1, 4):
            y = y0 + r * h / 3
            page.draw_line((x0, y), (x0 + w, y), color=(0.85, 0.88, 0.93))
        for c in range(1, 4):
            x = x0 + c * w / 4
            page.draw_line((x, y0), (x, y0 + h), color=(0.85, 0.88, 0.93))
        for r in range(3):
            for c in range(4):
                page.insert_text((x0 + c * w / 4 + 12, y0 + r * h / 3 + 26),
                                 f"P{i}-{r + 1}{c + 1}", fontsize=10, fontname="helv",
                                 color=(0.3, 0.35, 0.45))
        page.insert_text((58, 800), f"— 第 {i} 页 / 共 {pages} 页 —", fontsize=9, fontname="china-s",
                         color=(0.55, 0.6, 0.68))
    doc.save(str(path))
    doc.close()


def make_docx(path: Path, pages: int = 3) -> None:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Microsoft YaHei"
    style.font.size = Pt(10.5)

    for i in range(1, pages + 1):
        if i > 1:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        doc.add_heading(f"第 {i} 章 项目实施方案", level=1)
        doc.add_paragraph(
            f"这是 Word 样例的第 {i} 页。Word 自身没有固定页码，PagePull 通过 Office 把文档"
            "转换为 PDF 后再确定分页，因此这里的页码与 Word 打印预览完全一致。"
        )
        doc.add_heading("1.1 目标", level=2)
        for text in ("建立统一的资料提取流程", "减少手工截图与另存的工作量", "保证导出图片的清晰度一致"):
            doc.add_paragraph(text, style="List Bullet")

        table = doc.add_table(rows=4, cols=4)
        table.style = "Light Grid Accent 1"
        heads = ["序号", "工作项", "负责人", "完成时间"]
        for c, h in enumerate(heads):
            table.cell(0, c).text = h
        for r in range(1, 4):
            table.cell(r, 0).text = str(r)
            table.cell(r, 1).text = f"第 {i} 页任务 {r}"
            table.cell(r, 2).text = "张三"
            table.cell(r, 3).text = "2026-11-0%d" % r
        doc.add_paragraph(f"以上为第 {i} 页内容，用于验证 Word 按页导出图片的效果。")

    doc.save(str(path))


def make_xlsx(path: Path) -> None:
    wb = Workbook()
    thin = Side(style="thin", color="D0D7E2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws = wb.active
    ws.title = "销售明细"
    ws["A1"] = "2026 年第三季度销售明细"
    ws["A1"].font = Font(size=14, bold=True)
    ws.merge_cells("A1:H1")
    ws["A1"].alignment = Alignment(horizontal="center")
    heads = ["订单号", "日期", "客户", "产品", "数量", "单价", "金额", "状态"]
    for c, h in enumerate(heads, start=1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F6FED")
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
    for r in range(3, 33):
        ws.cell(row=r, column=1, value=f"SO2026{r:04d}")
        ws.cell(row=r, column=2, value=f"2026/9/{r - 1}")
        ws.cell(row=r, column=3, value=f"客户{r % 7 + 1}")
        ws.cell(row=r, column=4, value=["标准版", "专业版", "旗舰版"][r % 3])
        ws.cell(row=r, column=5, value=(r * 3) % 17 + 1)
        ws.cell(row=r, column=6, value=199.0 * (r % 3 + 1))
        ws.cell(row=r, column=7, value=f"=E{r}*F{r}")
        ws.cell(row=r, column=8, value="已回款" if r % 4 else "待回款")
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border
    for col, w in zip("ABCDEFGH", (14, 12, 12, 10, 8, 10, 12, 10)):
        ws.column_dimensions[col].width = w

    ws2 = wb.create_sheet("区域汇总")
    ws2["A1"] = "区域"
    ws2["B1"] = "金额"
    for i, (region, amount) in enumerate([("华东", 182000), ("华南", 143500), ("华北", 98200), ("西部", 61800)], start=2):
        ws2.cell(row=i, column=1, value=region)
        ws2.cell(row=i, column=2, value=amount)
    ws2.cell(row=6, column=1, value="合计").font = Font(bold=True)
    ws2.cell(row=6, column=2, value="=SUM(B2:B5)").font = Font(bold=True)

    ws3 = wb.create_sheet("使用说明")
    ws3["A1"] = "说明"
    ws3["A1"].font = Font(size=13, bold=True)
    for i, line in enumerate([
        "1. 本文件仅用于演示按工作表提取。",
        "2. 勾选「销售明细」导出 PNG，可得到整张表的高清图片。",
        "3. 勾选多张表并选择 Excel 格式，可生成新的工作簿。",
        "4. 宽表会按「一页宽」自动缩放后再渲染。",
    ], start=3):
        ws3.cell(row=i, column=1, value=line)
    ws3.column_dimensions["A"].width = 60

    wb.save(path)


def make_pptx(path: Path, slides: int = 4) -> None:
    """用本机 PowerPoint 现场生成一份演示文稿。

    没有 PowerPoint（也就没有 PowerPoint 的排版能力）时直接跳过，不影响其它样例。
    """
    try:
        import pythoncom
        import win32com.client as win32
    except Exception:
        print("跳过 PPT 样例：没有 pywin32")
        return

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32.DispatchEx("PowerPoint.Application")
        pres = app.Presentations.Add(0)       # WithWindow=False，不弹窗
        for i in range(1, slides + 1):
            layout = 1 if i == 1 else 2       # 1=标题, 2=标题+内容
            s = pres.Slides.Add(i, layout)
            if i == 1:
                s.Shapes(1).TextFrame.TextRange.Text = "小睦文件转换器 · PPT 样例"
                s.Shapes(2).TextFrame.TextRange.Text = (
                    "用于验证按幻灯片提取、渲染清晰度与导出格式"
                )
            else:
                s.Shapes(1).TextFrame.TextRange.Text = f"第 {i} 页：季度经营摘要"
                body = (
                    f"这是 PPT 样例的第 {i} 页。\n"
                    "· PPT 自身不能直接渲染，先由 PowerPoint 另存为 PDF 再取页；\n"
                    "· 因此导出的图片与 PowerPoint 里看到的完全一致；\n"
                    "· 勾选多页并选「PowerPoint 演示文稿」，可另存出一个子集 pptx。"
                )
                s.Shapes(2).TextFrame.TextRange.Text = body
        pres.SaveAs(str(path), 24)            # 24 = ppSaveAsOpenXMLPresentation
    except Exception as exc:  # noqa: BLE001
        print(f"跳过 PPT 样例：{exc}")
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    make_pdf(OUT / "样例报告.pdf")
    make_docx(OUT / "样例方案.docx")
    make_xlsx(OUT / "样例数据.xlsx")
    make_pptx(OUT / "样例演示.pptx")
    for p in sorted(OUT.iterdir()):
        print(f"{p.name:20s} {p.stat().st_size / 1024:8.1f} KB")


if __name__ == "__main__":
    main()
