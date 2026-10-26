"""无界面自测：解析样例文件，走完所有导出格式，并回读产物做校验。

运行： python tests/smoke_test.py
"""
from __future__ import annotations

import shutil
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pymupdf  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from PIL import Image  # noqa: E402

from core.exporter import run_jobs  # noqa: E402
from core.models import ExportSpec, Job  # noqa: E402
from core.sources import open_source  # noqa: E402

OUT = ROOT / "out_smoke"
SAMPLES = ROOT / "samples"

_results: list[tuple[bool, str, str]] = []


def check(label: str, fn) -> None:
    t0 = time.time()
    try:
        detail = fn() or ""
        _results.append((True, label, f"{detail} ({time.time() - t0:.1f}s)"))
    except Exception as exc:  # noqa: BLE001
        _results.append((False, label, f"{exc}"))
        traceback.print_exc()


def spec(fmt: str, **kw) -> ExportSpec:
    base = dict(out_dir=OUT / fmt, dpi=120, overwrite="overwrite")
    base.update(kw)
    return ExportSpec(fmt=fmt, **base)


def pick(src, keys: list[str]):
    units = [u for u in src.units() if u.key in set(keys)]
    assert len(units) == len(keys), f"缺少单位：{keys}"
    return Job(source=src, units=units)


def main() -> int:
    shutil.rmtree(OUT, ignore_errors=True)
    OUT.mkdir(parents=True, exist_ok=True)

    if not (SAMPLES / "样例报告.pdf").exists():
        print("缺少样例文件，先运行： python tools/make_samples.py")
        return 2

    # ---------------------------------------------------------------- PDF
    pdf_src = open_source(SAMPLES / "样例报告.pdf")
    print(f"PDF 单位数 = {len(pdf_src.units())}")

    def pdf_thumb() -> str:
        img = pdf_src.thumbnail("p1", 320)
        img.save(OUT / "preview_pdf.png")
        assert img.width > 100
        return f"{img.width}x{img.height}"

    def pdf_png() -> str:
        files = run_jobs([pick(pdf_src, ["p1", "p3", "p5"])], spec("png"))
        assert len(files) == 3, files
        for f in files:
            with Image.open(f) as im:
                assert im.width > 800 and im.height > 1000, (f, im.size)
        return f"{len(files)} 张图 {im.size[0]}x{im.size[1]}"

    def pdf_jpg_long() -> str:
        files = run_jobs([pick(pdf_src, ["p1", "p2", "p3"])], spec("jpg", long_image=True))
        assert len(files) == 1, files
        with Image.open(files[0]) as im:
            assert im.height > im.width * 2
        return f"长图 {im.width}x{im.height}"

    def pdf_pdf() -> str:
        files = run_jobs([pick(pdf_src, ["p1", "p2", "p5"])], spec("pdf"))
        doc = pymupdf.open(files[0])
        try:
            assert doc.page_count == 3, doc.page_count
            return f"{doc.page_count} 页"
        finally:
            doc.close()

    def pdf_docx() -> str:
        files = run_jobs([pick(pdf_src, ["p2", "p4"])], spec("docx", dpi=100))
        from docx import Document

        d = Document(str(files[0]))
        assert len(d.inline_shapes) == 2, len(d.inline_shapes)
        return f"{len(d.inline_shapes)} 张图已排入"

    # ---------------------------------------------------------------- Word
    word_src = open_source(SAMPLES / "样例方案.docx")
    print(f"Word 单位数 = {len(word_src.units())}")

    def word_png() -> str:
        files = run_jobs([pick(word_src, ["p1", "p2"])], spec("png"))
        with Image.open(files[0]) as im:
            return f"{len(files)} 张图 {im.size[0]}x{im.size[1]}"

    def word_pdf() -> str:
        files = run_jobs([pick(word_src, ["p1", "p3"])], spec("pdf"))
        doc = pymupdf.open(files[0])
        try:
            assert doc.page_count == 2
            text = doc[0].get_text()
            assert "实施方案" in text or len(text) > 20, text[:60]
            return f"{doc.page_count} 页，首页含文字 {len(text)} 字符"
        finally:
            doc.close()

    def word_docx() -> str:
        files = run_jobs([pick(word_src, ["p1"])], spec("docx", dpi=110))
        return f"{files[0].name} {files[0].stat().st_size // 1024} KB"

    # ---------------------------------------------------------------- Excel
    xls_src = open_source(SAMPLES / "样例数据.xlsx")
    print(f"Excel 单位数 = {len(xls_src.units())} -> {[u.label for u in xls_src.units()]}")

    def xls_thumb() -> str:
        img = xls_src.thumbnail("s::销售明细", 320)
        img.save(OUT / "preview_xlsx.png")
        return f"{img.width}x{img.height}"

    def xls_png() -> str:
        files = run_jobs([pick(xls_src, ["s::销售明细"])], spec("png", dpi=150))
        with Image.open(files[0]) as im:
            assert im.width > 800, im.size
            return f"{len(files)} 张图 {im.size[0]}x{im.size[1]}"

    def xls_xlsx() -> str:
        files = run_jobs([pick(xls_src, ["s::销售明细", "s::区域汇总"])], spec("xlsx"))
        wb = load_workbook(files[0])
        try:
            assert wb.sheetnames == ["销售明细", "区域汇总"], wb.sheetnames
            assert wb["销售明细"]["A2"].value == "订单号"
            return f"工作表 {wb.sheetnames}"
        finally:
            wb.close()

    def xls_xlsx_separate() -> str:
        files = run_jobs([pick(xls_src, ["s::区域汇总", "s::使用说明"])], spec("xlsx", merge_xlsx=False))
        assert len(files) == 2, files
        return f"{len(files)} 个工作簿"

    def xls_pdf() -> str:
        files = run_jobs([pick(xls_src, ["s::区域汇总"])], spec("pdf"))
        doc = pymupdf.open(files[0])
        try:
            assert doc.page_count >= 1
            return f"{doc.page_count} 页"
        finally:
            doc.close()

    check("PDF  缩略图", pdf_thumb)
    check("PDF  → PNG（非连续页 p1/p3/p5）", pdf_png)
    check("PDF  → JPG（三页拼长图）", pdf_jpg_long)
    check("PDF  → PDF（重排为 3 页）", pdf_pdf)
    check("PDF  → DOCX（图片按页排版）", pdf_docx)
    check("Word 缩略图 / 解析分页", lambda: f"{len(word_src.units())} 页")
    check("Word → PNG", word_png)
    check("Word → PDF（保留矢量文字）", word_pdf)
    check("Word → DOCX", word_docx)
    check("Excel 缩略图（工作表渲染）", xls_thumb)
    check("Excel → PNG（整表一图）", xls_png)
    check("Excel → XLSX（多表合并）", xls_xlsx)
    check("Excel → XLSX（一表一文件）", xls_xlsx_separate)
    check("Excel → PDF", xls_pdf)

    print("\n" + "=" * 64)
    ok = sum(1 for r in _results if r[0])
    for passed, label, detail in _results:
        print(f"{'✔ PASS' if passed else '✘ FAIL'}  {label:34s} {detail}")
    print("=" * 64)
    print(f"合计 {ok}/{len(_results)} 通过；产物目录：{OUT}")
    total = sum(f.stat().st_size for f in OUT.rglob("*") if f.is_file())
    print(f"产物 {sum(1 for f in OUT.rglob('*') if f.is_file())} 个文件，{total / 1024:.0f} KB")
    return 0 if ok == len(_results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
