"""Office / LibreOffice 桥接层。

Word 与 Excel 没有"页"或"可直接渲染"的概念，高保真做法是借用本机 Office：
  Word  -> 导出 PDF（分页由此确定）-> 再由 PyMuPDF 渲染/抽页
  Excel -> 单表复制到临时工作簿 -> 导出 PDF / 另存 xlsx
无 Office 时降级：Excel 用 openpyxl 拷贝单元格；Word 无法确定页码，直接报错。
"""
from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import tempfile
import threading
import time
from pathlib import Path
from typing import Iterable

_CACHE_ROOT = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "PagePull" / "cache"
_COM_LOCK = threading.RLock()          # Office COM 不适合多线程并发
_CAPS: dict[str, bool | str | None] = {}


# --------------------------------------------------------------------------- 基础

def _tag_of(*parts) -> str:
    """稳定的缓存键：不能用内置 hash()（进程间随机化，会导致跨会话缓存失效）。"""
    h = hashlib.md5()
    for p in parts:
        h.update(str(p).encode("utf-8", "ignore"))
    return h.hexdigest()[:12]


def cache_dir(*parts: str) -> Path:
    d = _CACHE_ROOT.joinpath(*parts)
    d.mkdir(parents=True, exist_ok=True)
    return d


def _have_pywin32() -> bool:
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
        return True
    except Exception:
        return False


def soffice_path() -> str | None:
    if _CAPS.get("soffice") is not None:
        return _CAPS["soffice"] or None  # type: ignore[return-value]
    cand = shutil.which("soffice") or shutil.which("soffice.exe")
    if not cand:
        for p in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if Path(p).exists():
                cand = p
                break
    _CAPS["soffice"] = cand or ""
    return cand


def has_word() -> bool:
    return _probe("word", "Word.Application")


def has_excel() -> bool:
    return _probe("excel", "Excel.Application")


def _probe(key: str, progid: str) -> bool:
    if key in _CAPS:
        return bool(_CAPS[key])
    ok = False
    if _have_pywin32():
        try:
            import pythoncom
            import win32com.client as win32

            with _COM_LOCK:
                pythoncom.CoInitialize()
                try:
                    app = win32.DispatchEx(progid)
                    try:
                        _ = app.Version
                    finally:
                        app.Quit()
                    ok = True
                finally:
                    pythoncom.CoUninitialize()
        except Exception:
            ok = False
    _CAPS[key] = ok
    return ok


def capabilities() -> dict[str, bool]:
    return {"word": has_word(), "excel": has_excel(), "soffice": bool(soffice_path())}


# --------------------------------------------------------------------------- Word

def word_to_pdf(src: Path, *, force: bool = False) -> Path:
    """把 Word 文档转成 PDF，按 (路径, 大小, 修改时间) 缓存，避免重复转换。"""
    src = Path(src)
    st = src.stat()
    tag = _tag_of(str(src).lower(), st.st_size, int(st.st_mtime))
    out = cache_dir("word") / f"{src.stem}_{tag}.pdf"
    if out.exists() and out.stat().st_size > 0 and not force:
        return out
    if has_word():
        _word_export(src, out)
    elif soffice_path():
        _soffice_convert(src, out, "pdf")
    else:
        raise RuntimeError("未检测到 Microsoft Word 或 LibreOffice，无法确定 Word 文档的页码。")
    if not out.exists() or out.stat().st_size == 0:
        raise RuntimeError("Word 转 PDF 失败，输出为空。")
    return out


def _word_export(src: Path, out: Path) -> None:
    import pythoncom
    import win32com.client as win32

    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = None
        doc = None
        try:
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            app.Options.UpdateLinksAtOpen = False
            doc = app.Documents.Open(
                str(src), False, True, False, "", "", False, "", "", 0, 0, False
            )  # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ..., Visible
            try:
                doc.Repaginate()
            except Exception:
                pass
            doc.ExportAsFixedFormat(
                str(out),
                17,       # ExportFormat：wdExportFormatPDF
                False,    # OpenAfterExport
                0,        # OptimizeFor：打印质量
                0,        # Range：整个文档
                0,        # From
                0,        # To
                0,        # Item：正文
                True,     # IncludeDocProps
                True,     # KeepIRM
                0,        # CreateBookmarks
                True,     # DocStructureTags
                True,     # BitmapMissingFonts
                False,    # UseISO19005_1
            )
        finally:
            # 先关文档再退进程，避免用户已打开的 Word 被牵动
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit(SaveChanges=0)
                except Exception:
                    pass
            pythoncom.CoUninitialize()


def _soffice_convert(src: Path, out_pdf: Path, kind: str) -> None:
    exe = soffice_path()
    if not exe:
        raise RuntimeError("未检测到 LibreOffice。")
    tmpdir = Path(tempfile.mkdtemp(prefix="pp_lo_"))
    cmd = [exe, "--headless", "--norestore", "--convert-to", kind, "--outdir", str(tmpdir), str(src)]
    subprocess.run(cmd, check=True, timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    produced = list(tmpdir.glob(f"*.{kind}"))
    if not produced:
        raise RuntimeError("LibreOffice 转换未产出文件。")
    shutil.move(str(produced[0]), str(out_pdf))
    shutil.rmtree(tmpdir, ignore_errors=True)


# --------------------------------------------------------------------------- Excel

def excel_sheets(path: Path) -> list[dict]:
    """返回 [{'name','visible','rows','cols'}]。优先 COM（能拿到真实使用区域），否则 openpyxl。"""
    path = Path(path)
    if has_excel() and path.suffix.lower() not in (".xlsb",):
        try:
            return _excel_sheets_com(path)
        except Exception:
            pass
    return _excel_sheets_openpyxl(path)


def _excel_sheets_com(path: Path) -> list[dict]:
    import pythoncom
    import win32com.client as win32

    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = wb = None
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(str(path), 0, True)  # FileName, UpdateLinks=0, ReadOnly=True
            out: list[dict] = []
            for ws in wb.Worksheets:
                try:
                    used = ws.UsedRange
                    rows = int(used.Rows.Count)
                    cols = int(used.Columns.Count)
                except Exception:
                    rows = cols = 0
                out.append({
                    "name": str(ws.Name),
                    "visible": int(ws.Visible) == -1,
                    "rows": rows,
                    "cols": cols,
                })
            return out
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


def _excel_sheets_openpyxl(path: Path) -> list[dict]:
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        out = []
        for ws in wb.worksheets:
            out.append({
                "name": ws.title,
                "visible": bool(ws.sheet_state == "visible"),
                "rows": int(ws.max_row or 0),
                "cols": int(ws.max_column or 0),
            })
        return out
    finally:
        wb.close()


def excel_sheet_to_pdf(path: Path, sheet: str, *, fit_one_page_wide: bool = True) -> Path:
    """把单张工作表导出为 PDF（用于预览与图片导出）。"""
    path = Path(path)
    st = path.stat()
    tag = _tag_of(str(path).lower(), sheet, st.st_size, int(st.st_mtime), fit_one_page_wide)
    out = cache_dir("excel") / f"{path.stem}_{_tag_of(sheet)}_{tag}.pdf"
    if out.exists() and out.stat().st_size > 0:
        return out
    if has_excel():
        try:
            _excel_sheet_export_pdf(path, sheet, out, fit_one_page_wide)
            return out
        except Exception as exc:
            raise RuntimeError(f"Excel 导出工作表「{sheet}」失败：{exc}") from exc
    if soffice_path():
        _soffice_convert(path, out, "pdf")   # 注意：整簿转 PDF，仅首表可用
        raise RuntimeError("未检测到 Microsoft Excel，仅支持工作表另存为 xlsx，暂不能渲染为图片。")
    raise RuntimeError("未检测到 Microsoft Excel，无法把工作表渲染为图片。")


def _excel_sheet_export_pdf(path: Path, sheet: str, out: Path, fit_one_page_wide: bool) -> None:
    import pythoncom
    import win32com.client as win32

    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = src = new_wb = None
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            src = app.Workbooks.Open(str(path), 0, True)  # UpdateLinks=0, ReadOnly=True
            src.Worksheets(sheet).Copy()           # 单表复制 -> 新工作簿
            new_wb = app.ActiveWorkbook
            ws = new_wb.Worksheets(1)
            if fit_one_page_wide:
                try:
                    ws.PageSetup.Zoom = False
                    ws.PageSetup.FitToPagesWide = 1
                    ws.PageSetup.FitToPagesTall = False
                    ws.PageSetup.Orientation = 2 if (ws.UsedRange.Columns.Count > ws.UsedRange.Rows.Count) else 1
                except Exception:
                    pass
            new_wb.ExportAsFixedFormat(0, str(out))  # 0 = xlTypePDF
        finally:
            if new_wb is not None:
                try:
                    new_wb.Close(SaveChanges=False)
                except Exception:
                    pass
            if src is not None:
                try:
                    src.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


def excel_extract_sheets(path: Path, sheets: Iterable[str], out: Path) -> None:
    """把若干工作表另存为一个新工作簿，保留格式/图表/列宽。"""
    names = [s for s in sheets]
    if not names:
        raise ValueError("未选择工作表。")
    if has_excel():
        try:
            _excel_extract_com(path, names, out)
            return
        except Exception as exc:
            if len(names) > 1:
                raise RuntimeError(f"Excel 复制工作表失败：{exc}") from exc
    _excel_extract_openpyxl(path, names, out)


def _excel_extract_com(path: Path, names: list[str], out: Path) -> None:
    import pythoncom
    import win32com.client as win32

    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = src = new_wb = None
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            src = app.Workbooks.Open(str(path), 0, True)  # UpdateLinks=0, ReadOnly=True
            src.Worksheets(names[0]).Copy()
            new_wb = app.ActiveWorkbook
            # 注意：IDispatch 动态派发下关键字参数会被静默丢弃，必须用位置参数 Copy(Before, After)
            for extra in names[1:]:
                src.Worksheets(extra).Copy(None, new_wb.Worksheets(new_wb.Worksheets.Count))
            new_wb.SaveAs(str(out), 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
        finally:
            if new_wb is not None:
                try:
                    new_wb.Close(SaveChanges=False)
                except Exception:
                    pass
            if src is not None:
                try:
                    src.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


def _excel_extract_openpyxl(path: Path, names: list[str], out: Path) -> None:
    """降级方案：只搬单元格值与基础格式，图片/图表/条件格式会丢失。"""
    from copy import copy

    from openpyxl import Workbook, load_workbook

    src = load_workbook(path, data_only=False)
    dst = Workbook()
    dst.remove(dst.active)
    for name in names:
        if name not in src.sheetnames:
            continue
        s = src[name]
        d = dst.create_sheet(title=name[:31])
        for row in s.iter_rows():
            for cell in row:
                if cell.value is None and not cell.has_style:
                    continue
                t = d.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    t.font = copy(cell.font)
                    t.fill = copy(cell.fill)
                    t.border = copy(cell.border)
                    t.alignment = copy(cell.alignment)
                    t.number_format = cell.number_format
        for rng in s.merged_cells.ranges:
            try:
                d.merge_cells(str(rng))
            except Exception:
                pass
        for col, dim in s.column_dimensions.items():
            if dim.width:
                d.column_dimensions[col].width = dim.width
        for idx, dim in s.row_dimensions.items():
            if dim.height:
                d.row_dimensions[idx].height = dim.height
        d.freeze_panes = s.freeze_panes
    dst.save(out)
    src.close()


def cleanup_stale(max_age_hours: int = 72) -> None:
    """清掉过期的缓存 PDF/图片。"""
    if not _CACHE_ROOT.exists():
        return
    cutoff = time.time() - max_age_hours * 3600
    for p in _CACHE_ROOT.rglob("*"):
        try:
            if p.is_file() and p.stat().st_mtime < cutoff:
                p.unlink()
        except Exception:
            pass
