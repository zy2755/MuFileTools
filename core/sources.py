"""四种来源的统一封装：PDF / Word / PPT / Excel。

对外只暴露一套接口：
    open()                     解析文件
    units()                    可勾选的单位列表（页 或 工作表）
    thumbnail(key, width)      预览缩略图 -> PIL.Image
    render(key, dpi)           高清渲染 -> [PIL.Image]（工作表可能跨多页）
    export_pdf(keys, out)      原生抽页 / 合并为 PDF
    export_workbook(keys, out) Excel 专用：另存工作簿
    export_presentation(...)   PPT 专用：另存演示文稿
"""
from __future__ import annotations

from pathlib import Path

import pymupdf
from PIL import Image

from . import office
from .models import Kind, Unit


def _pil(pix) -> Image.Image:
    mode = "RGBA" if pix.alpha else "RGB"
    return Image.frombytes(mode, (pix.width, pix.height), pix.samples)


class PdfSource:
    kind: Kind = "pdf"

    def __init__(self, path: Path):
        self.path = Path(path)
        self._doc: pymupdf.Document | None = None
        self.page_count = 0

    # ------------------------------------------------------------------ 生命周期
    def open(self, progress=None) -> None:
        self._doc = pymupdf.open(self.path)
        self.page_count = self._doc.page_count
        if self.page_count == 0:
            raise RuntimeError("该 PDF 没有任何页面。")

    def close(self) -> None:
        if self._doc is not None:
            try:
                self._doc.close()
            except Exception:
                pass
            self._doc = None

    @property
    def doc(self) -> pymupdf.Document:
        if self._doc is None:
            raise RuntimeError("文档尚未打开。")
        return self._doc

    # ------------------------------------------------------------------ 单位
    def units(self) -> list[Unit]:
        out = []
        for i in range(self.page_count):
            page = self.doc[i]
            w, h = page.rect.width, page.rect.height
            out.append(Unit(key=f"p{i + 1}", label=f"第 {i + 1} 页",
                            meta={"index": i, "size": (round(w), round(h))}))
        return out

    # ------------------------------------------------------------------ 渲染
    def thumbnail(self, key: str, width: int = 240) -> Image.Image:
        page = self.doc[_index(key)]
        w = page.rect.width or 1
        zoom = max(0.05, min(width / w, 4.0))
        return _pil(page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False))

    def render(self, key: str, dpi: int = 200) -> list[Image.Image]:
        page = self.doc[_index(key)]
        zoom = dpi / 72.0
        return [_pil(page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False))]

    # ------------------------------------------------------------------ 导出
    def export_pdf(self, keys: list[str], out_path: Path) -> None:
        order = sorted({_index(k) for k in keys})
        if not order:
            raise ValueError("未选择任何页面。")
        target = pymupdf.open()
        try:
            for i in order:
                target.insert_pdf(self.doc, from_page=i, to_page=i, links=True, annots=True)
            target.save(str(out_path), garbage=3, deflate=True)
        finally:
            target.close()


class WordSource:
    """Word 靠"先转 PDF"获得稳定分页，之后完全复用 PDF 能力。"""

    kind: Kind = "word"

    def __init__(self, path: Path):
        self.path = Path(path)
        self._pdf: PdfSource | None = None

    def open(self, progress=None) -> None:
        if progress:
            progress("正在用 Office 解析 Word 分页…")
        pdf_path = office.word_to_pdf(self.path)
        self._pdf = PdfSource(pdf_path)
        self._pdf.open()

    def close(self) -> None:
        if self._pdf:
            self._pdf.close()

    @property
    def doc(self) -> PdfSource:
        if self._pdf is None:
            raise RuntimeError("文档尚未打开。")
        return self._pdf

    @property
    def page_count(self) -> int:
        return self.doc.page_count

    def units(self) -> list[Unit]:
        out = []
        for u in self.doc.units():
            out.append(Unit(key=u.key, label=u.label, meta={**u.meta, "from": "word"}))
        return out

    def thumbnail(self, key: str, width: int = 240) -> Image.Image:
        return self.doc.thumbnail(key, width)

    def render(self, key: str, dpi: int = 200) -> list[Image.Image]:
        return self.doc.render(key, dpi)

    def export_pdf(self, keys: list[str], out_path: Path) -> None:
        self.doc.export_pdf(keys, out_path)


class PptSource:
    """PPT 和 Word 一个路子：先转 PDF 拿到"每页幻灯片"，之后完全复用 PDF 能力。

    额外多一个 export_presentation()，用于把所选幻灯片另存为新的 .pptx。
    """

    kind: Kind = "ppt"

    def __init__(self, path: Path):
        self.path = Path(path)
        self._pdf: PdfSource | None = None

    def open(self, progress=None) -> None:
        if progress:
            progress("正在用 PowerPoint 解析幻灯片…")
        self._pdf = PdfSource(office.ppt_to_pdf(self.path))
        self._pdf.open()

    def close(self) -> None:
        if self._pdf:
            self._pdf.close()

    @property
    def doc(self) -> PdfSource:
        if self._pdf is None:
            raise RuntimeError("文档尚未打开。")
        return self._pdf

    @property
    def page_count(self) -> int:
        return self.doc.page_count

    def units(self) -> list[Unit]:
        out = []
        for u in self.doc.units():
            out.append(Unit(key=u.key, label=u.label, meta={**u.meta, "from": "ppt"}))
        return out

    def thumbnail(self, key: str, width: int = 240) -> Image.Image:
        return self.doc.thumbnail(key, width)

    def render(self, key: str, dpi: int = 200) -> list[Image.Image]:
        return self.doc.render(key, dpi)

    def export_pdf(self, keys: list[str], out_path: Path) -> None:
        self.doc.export_pdf(keys, out_path)

    def export_presentation(self, keys: list[str], out_path: Path) -> None:
        """把所选幻灯片另存为一个新演示文稿（页键 pN 是 1 起，PowerPoint 也是 1 起）。"""
        office.ppt_extract_slides(self.path, [_index(k) + 1 for k in keys], out_path)


class ExcelSource:
    """Excel 以"工作表"为单位；渲染借 Excel 打印为 PDF 再转图。"""

    kind: Kind = "excel"

    def __init__(self, path: Path):
        self.path = Path(path)
        self._sheets: list[dict] = []
        self._pdf_cache: dict[str, Path] = {}

    def open(self, progress=None) -> None:
        if progress:
            progress("正在读取工作表结构…")
        self._sheets = office.excel_sheets(self.path)
        if not self._sheets:
            raise RuntimeError("该工作簿没有工作表。")

    def close(self) -> None:
        self._pdf_cache.clear()

    @property
    def page_count(self) -> int:
        return len(self._sheets)

    def units(self) -> list[Unit]:
        return [
            Unit(
                key=f"s::{s['name']}",
                label=s["name"],
                meta={"rows": s["rows"], "cols": s["cols"], "visible": s["visible"]},
            )
            for s in self._sheets
        ]

    # ------------------------------------------------------------------ 渲染
    def _sheet_pdf(self, key: str) -> Path:
        name = _sheet_name(key)
        if name not in self._pdf_cache:
            self._pdf_cache[name] = office.excel_sheet_to_pdf(self.path, name)
        return self._pdf_cache[name]

    def thumbnail(self, key: str, width: int = 240) -> Image.Image:
        pdf = PdfSource(self._sheet_pdf(key))
        pdf.open()
        try:
            return pdf.thumbnail("p1", width)
        finally:
            pdf.close()

    def render(self, key: str, dpi: int = 200) -> list[Image.Image]:
        pdf = PdfSource(self._sheet_pdf(key))
        pdf.open()
        try:
            return [pdf.render(f"p{i + 1}", dpi)[0] for i in range(pdf.page_count)]
        finally:
            pdf.close()

    def export_pdf(self, keys: list[str], out_path: Path) -> None:
        target = pymupdf.open()
        try:
            for k in keys:
                src = pymupdf.open(self._sheet_pdf(k))
                try:
                    target.insert_pdf(src)
                finally:
                    src.close()
            if target.page_count == 0:
                raise ValueError("未选择任何工作表。")
            target.save(str(out_path), garbage=3, deflate=True)
        finally:
            target.close()

    def export_workbook(self, keys: list[str], out_path: Path) -> None:
        office.excel_extract_sheets(self.path, [_sheet_name(k) for k in keys], out_path)


# --------------------------------------------------------------------------- 工厂

def open_source(path: str | Path, progress=None):
    from .models import SUPPORTED_EXTS

    p = Path(path)
    kind = SUPPORTED_EXTS.get(p.suffix.lower())
    if kind is None:
        raise RuntimeError(f"暂不支持的文件类型：{p.suffix}")
    if not p.exists():
        raise RuntimeError("文件不存在。")
    cls = {"pdf": PdfSource, "word": WordSource, "excel": ExcelSource, "ppt": PptSource}[kind]
    src = cls(p)
    src.open(progress=progress)
    return src


def _index(key: str) -> int:
    try:
        return int(key[1:]) - 1
    except Exception as exc:
        raise ValueError(f"非法页键：{key}") from exc


def _sheet_name(key: str) -> str:
    if not key.startswith("s::"):
        raise ValueError(f"非法工作表键：{key}")
    return key[3:]
