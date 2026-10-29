"""Office / LibreOffice 桥接层。

Word / PPT / Excel 都没有"可直接渲染"的概念，高保真做法是借用本机 Office：
  Word -> 导出 PDF（分页由此确定）-> 再由 PyMuPDF 渲染/抽页
  PPT  -> 导出 PDF（一页幻灯片 = 一页 PDF）-> 同上；另可"删页另存"出子集 pptx
  Excel-> 单表复制到临时工作簿 -> 导出 PDF / 另存 xlsx
无 Office 时降级：Excel 用 openpyxl 拷贝单元格；Word / PPT 无法确定分页，直接报错。
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


def has_powerpoint() -> bool:
    """PowerPoint 的探测不能真去启动它：它的主窗口不允许隐藏
    （Application.Visible = False 会报 "Hiding the application window is not allowed."），
    一探就会在屏幕上闪一个灰窗口。所以先查注册表，注册了就当有。"""
    if key_registered("PowerPoint.Application"):
        _CAPS["ppt"] = True
        return True
    return _probe("ppt", "PowerPoint.Application")


def key_registered(progid: str) -> bool:
    """COM 组件是否注册（只读注册表，不启动任何程序）。"""
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{progid}\CurVer"):
            return True
    except OSError:
        return False
    except Exception:
        return False


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
                        # 别忘了先藏起来：DispatchEx 建出来的应用默认是可见的，
                        # 直接读 Version 会在屏幕正中闪一个空白窗口（Word 是"文档1"）。
                        app.Visible = False
                    except Exception:
                        pass
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
    return {
        "word": has_word(),
        "excel": has_excel(),
        "ppt": has_powerpoint(),
        "soffice": bool(soffice_path()),
    }


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


# --------------------------------------------------------------------------- Word 合并

# Word 的常量（动态派发下只能按位置传数值）
_WD_GOTO_PAGE = 1
_WD_GOTO_ABSOLUTE = 1
_WD_STAT_PAGES = 2
_WD_PAGE_BREAK = 7
_WD_COLLAPSE_END = 0
_WD_FORMAT_DOCX = 16
_WD_BREAK_CHAR = "\x0c"     # 手动分页符在 Range.Text 里的样子


def _wd_page_starts(doc) -> list[int]:
    """每页正文的起始偏移，末尾补一个「全文结束」偏移（长度 = 页数 + 1）。

    注意：GoTo 按页定位**必须有文档窗口**。用 Open(..., Visible=False) 打开的文档
    没有窗口，GoTo 会让 WINWORD 直接崩掉（报「远程过程调用失败」，进程消失）。
    所以打开时最后一个参数 Visible 要传 True —— 应用本身仍是隐藏的，不会闪窗。
    """
    doc.Repaginate()
    total = int(doc.ComputeStatistics(_WD_STAT_PAGES))
    starts = [int(doc.GoTo(_WD_GOTO_PAGE, _WD_GOTO_ABSOLUTE, i).Start) for i in range(1, total + 1)]
    starts.append(int(doc.Content.End))
    return starts


def _wd_page_span(doc, starts: list[int], page: int) -> tuple[int, int, bool]:
    """取第 page 页（1 起）的范围，并剥掉结尾自带的分页符。

    返回 (start, end, closed)：closed 表示这段内容结尾已经是一个完整段落。
    剥不剥很关键 —— 不剥的话每页都会带着原来的分页符，再叠上我们自己插的分页符，
    一页内容会变成两页（实测 4 页输出成 8 页就是这么来的）。
    """
    if not 1 <= page < len(starts):
        raise ValueError(f"原文只有 {len(starts) - 1} 页，取不到第 {page} 页。")
    start, end = starts[page - 1], starts[page]
    text = doc.Range(start, end).Text
    if text.endswith(_WD_BREAK_CHAR + "\r"):
        end -= 2
        closed = text[:-2].endswith("\r")
    elif text.endswith(_WD_BREAK_CHAR):
        end -= 1
        closed = text[:-1].endswith("\r")
    else:
        closed = text.endswith("\r")
    return start, max(start, end), closed


def _wd_copy_page_setup(src_doc, dst_doc) -> None:
    """新文档的纸张/页边距照抄第一个来源，否则 A4 的内容会按 Letter 重新排版。"""
    try:
        s, d = src_doc.PageSetup, dst_doc.PageSetup
    except Exception:
        return
    for attr in (
        "PageWidth", "PageHeight", "TopMargin", "BottomMargin",
        "LeftMargin", "RightMargin", "Gutter", "Orientation",
        "HeaderDistance", "FooterDistance",
    ):
        try:
            setattr(d, attr, getattr(s, attr))
        except Exception:
            pass


def merge_docx(segments, out_path: Path, *, dpi: int = 200, progress=None,
               toc_levels: int | None = None) -> None:
    """把若干「页」原生拼成一份新的 .docx。

    segments 按输出顺序排列，每个元素二选一：
        ("word", Path, page_no)   Word 来源：按页取 Range，文字可选可编辑
        ("image", Path)           其它来源（PDF/PPT/Excel）：已经渲染好的整页 PNG

    toc_levels 非空时，在保存前顺手插一个目录（1..4 = 收录到第几级标题）。
    这里复用同一个 Word 会话，比"合并完再打开一次插目录"省一次冷启动。

    会打开本机 Word（隐藏），同一个会话里把所有源文档打开一次，逐页拼装。
    页眉页脚不在 Range 里，合并后统一用新文档的（一般就是没有）—— 这点在界面上说明了。
    """
    import pythoncom
    import win32com.client as win32

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink()

    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = new_doc = None
        sources: dict[str, object] = {}
        starts_cache: dict[str, list[int]] = {}
        try:
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            try:
                app.Options.UpdateLinksAtOpen = False
            except Exception:
                pass

            def open_src(path: Path):
                key = str(path).lower()
                if key not in sources:
                    # 最后一个参数 Visible=True：给文档建出窗口，GoTo 才不崩
                    doc = app.Documents.Open(
                        str(path), False, True, False, "", "", False, "", "", 0, 0, True
                    )
                    sources[key] = doc
                    starts_cache[key] = _wd_page_starts(doc)
                return sources[key], starts_cache[key]

            # 先把所有 Word 源打开（顺便预热分页），再建目标文档
            for seg in segments:
                if seg[0] != "word":
                    continue
                open_src(Path(seg[1]))

            new_doc = app.Documents.Add()
            if sources:
                _wd_copy_page_setup(next(iter(sources.values())), new_doc)
            try:
                usable = float(new_doc.PageSetup.PageWidth) - float(new_doc.PageSetup.LeftMargin) \
                    - float(new_doc.PageSetup.RightMargin)
            except Exception:
                usable = 0.0

            first = True
            total = len(segments)
            for i, seg in enumerate(segments, start=1):
                if progress:
                    progress(i, total, f"拼装第 {i}/{total} 页")
                if not first:
                    tail = new_doc.Content
                    tail.Collapse(_WD_COLLAPSE_END)
                    tail.InsertBreak(_WD_PAGE_BREAK)
                first = False

                dst = new_doc.Content
                dst.Collapse(_WD_COLLAPSE_END)

                if seg[0] == "word":
                    doc, starts = open_src(Path(seg[1]))
                    a, b, closed = _wd_page_span(doc, starts, int(seg[2]))
                    dst.FormattedText = doc.Range(a, b).FormattedText
                    if not closed:
                        # 页边界切在段落中间：补一个段落标记，免得下一页的内容接在同一段里
                        tail = new_doc.Content
                        tail.Collapse(_WD_COLLAPSE_END)
                        tail.InsertAfter("\r")
                else:
                    pic = Path(seg[1])
                    shape = new_doc.InlineShapes.AddPicture(str(pic), False, True, dst)
                    if usable > 0:
                        try:
                            shape.LockAspectRatio = -1
                            shape.Width = usable
                        except Exception:
                            pass

            if toc_levels:
                if progress:
                    progress(total, total, "生成目录…")
                _wd_insert_toc_into(new_doc, int(toc_levels))

            new_doc.SaveAs2(str(out_path), _WD_FORMAT_DOCX)
        finally:
            if new_doc is not None:
                try:
                    new_doc.Close(SaveChanges=0)
                except Exception:
                    pass
            for doc in sources.values():
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


# --------------------------------------------------------------------------- Word 目录

# 动态派发下只能按位置传数值，所以这些常量必须写成数字
_WD_ON = -1                     # ParagraphFormat.PageBreakBefore = True
_WD_ALIGN_CENTER = 1            # wdAlignParagraphCenter
_TOC_TITLE = "目录"
_COVER_STYLES = ("标题", "Title", "副标题", "Subtitle")
_TOC_EMPTY_HINTS = ("未找到目录项", "No table of contents entries")


def _wd_para_text(p) -> str:
    """取段落纯文本（去掉段落标记 \\r 和单元格标记 \\x07）。"""
    try:
        raw = str(p.Range.Text)
    except Exception:
        return ""
    return raw.replace("\r", "").replace("\x07", "").strip()


def _wd_cover_field_count(doc, from_idx: int, *, span: int = 20) -> int:
    """从第 from_idx 段往后数「封面字段行」有几条（够 2 条就提前返回）。

    用途：判断一个候选标题是不是**还在封面上**。封面大标题和章标题在 Word 眼里
    没有区别（都是大纲级别 1），但封面标题后面一定跟着一串
    「学院：」「姓名：」「指导教师：」的报名表式字段 —— 正文标题后面不会有。
    """
    try:
        total = int(doc.Paragraphs.Count)
    except Exception:
        return 0
    from .outline import _COVER_FIELD_RE      # 正则在 outline 里，两处共用一份
  
    count = 0
    for j in range(from_idx + 1, min(total, from_idx + span) + 1):
        try:
            text = _wd_para_text(doc.Paragraphs(j))
        except Exception:
            break
        if text and _COVER_FIELD_RE.match(text):
            count += 1
            if count >= 2:
                return count
    return count


def _wd_first_heading_pos(doc, *, scan: int = 400) -> int:
    """返回第一个「正文标题」段落的起始字符位置；找不到返回 0。

    判据用 OutlineLevel 而不是样式名：`wdOutlineLevel1..9` 是 1..9，正文是 10
    （`wdOutlineLevelBodyText`）。这样「标题 1 / Heading 1 / 自定义大纲级别」都能认，
    也不用去猜中文还是英文的样式名。

    为什么用它当插入点，而不是量「封面块有多长」：
    封面经常是手打的（正文样式 + 大字），根本没有 Title/Subtitle 样式，
    按样式量长度会得到 0，于是目录被插到封面**前面**去（真实踩过）。
    而「第一个正文标题之前」这个位置，无论封面多长、有没有样式，都天然是正文开头。
    目录条目自己的样式是 TOC 1/2/3，OutlineLevel 是正文值，不会被误认成标题。

    三个「不算标题」的例外，每一条都是真踩出来的：

    * **空段落**。WPS 存出来的文档开头常常躺着一个空的「标题 3」段落（格式残留），
      它的 OutlineLevel 是 3，按老逻辑第一个就命中它、返回位置 0，
      于是目录又被插回封面之前 —— v1.6.0 收到的那份调查报告就是这么错的。
    * **表格里的段落**。封面信息常用表格排版，格子里套着标题样式也不该当插入点。
    * **目录域自己的行**。样式是 TOC 1..9，大纲级别本来是正文值，但被用户手动标过
      就会变成 1..9。跳过 `TOC` 开头的样式，避免把新目录插进旧目录中间。
    """
    try:
        count = int(doc.Paragraphs.Count)
    except Exception:
        return 0
    for i in range(1, min(count, scan) + 1):
        try:
            p = doc.Paragraphs(i)
            level = int(p.OutlineLevel)
        except Exception:
            break
        if not (1 <= level <= 9):
            continue
        if not _wd_para_text(p):
            continue                      # 空段落不是标题
        try:
            if p.Range.Information(12):   # wdWithInTable
                continue
        except Exception:
            pass
        try:
            if _wd_is_toc_style_para(p):
                continue                  # 旧目录的条目行
        except Exception:
            pass
        if _wd_cover_field_count(doc, i) >= 2:
            continue                      # 后面跟着一串报名表字段 → 这是封面标题
        return int(p.Range.Start)
    return 0


def _wd_is_toc_style_para(para) -> bool:
    """段落样式是不是目录条目（TOC 1..9 / 目录 1..9 / TOC Heading）。"""
    try:
        name = str(para.Style.NameLocal).upper()
    except Exception:
        return False
    if name.startswith("TOC"):
        return True
    return name.startswith("目录") and name[-1:].isdigit()


def _wd_flatten_empty_headings(doc) -> int:
    """把「空段落」身上的标题属性清掉，返回处理条数。

    空的大纲段落会被目录收成一条**只有页码、没有标题字**的怪条目
    （实测：目录第一行孤零零一个「2」）。它不承载任何内容，
    清掉对排版几乎零影响，只是让它别进目录。

    **必须改样式，不能改 OutlineLevel**：实测对「标题 1/2/3」这类内置样式段落，
    `p.OutlineLevel = 10` 和 `p.Range.ParagraphFormat.OutlineLevel = 10`
    两种写法都**静默失效**（赋值不报错，读回还是 3）—— Word 认为大纲级别
    由标题样式决定，不接受段落级覆盖。改套「正文」样式才真的生效。
    """
    try:
        count = int(doc.Paragraphs.Count)
    except Exception:
        return 0
    done = 0
    for i in range(1, count + 1):
        try:
            p = doc.Paragraphs(i)
            if not (1 <= int(p.OutlineLevel) <= 9):
                continue
            if _wd_para_text(p):
                continue
            p.Style = -1                   # wdStyleNormal：空段落套标题样式没有意义
            done += 1
        except Exception:
            continue
    return done


def _wd_para_index_at(doc, pos: int) -> int:
    """字符位置 pos 落在第几段（Word 的 1 基下标）；找不到返回 0。"""
    try:
        count = int(doc.Paragraphs.Count)
    except Exception:
        return 0
    for i in range(1, count + 1):
        try:
            rng = doc.Paragraphs(i).Range
            if int(rng.Start) <= pos <= int(rng.End):
                return i
        except Exception:
            return 0
    return 0


def _wd_back_over_empty(doc, pos: int) -> int:
    """把插入点往前挪过**紧邻的空段落**，返回新的字符位置。

    封面和正文之间常夹着一堆空段落（WPS 排版残留）。目录插在它们**之后**，
    再给目录设「段前分页」，那些空段落就被单独留在一页上 —— 实测封面第 1 页、
    空页第 2 页、目录第 3 页。往前挪到空段落之前，空页就没了。

    只跳过真正的空段落（无文字、无图、无表格标记），遇到有内容的立刻停。
    """
    idx = _wd_para_index_at(doc, pos)
    if idx <= 1:
        return pos
    j = idx - 1
    while j >= 1:
        try:
            p = doc.Paragraphs(j)
            if _wd_para_text(p):
                break
            pos = int(p.Range.Start)
        except Exception:
            break
        j -= 1
    return pos


def _wd_next_nonempty_pos(doc, pos: int) -> int:
    """从 pos 往后找第一个非空段落的起点，返回它；找不到返回 pos 本身。

    用来把「正文另起一页」的分页属性放在**正文第一段**上，
    而不是目录后面那个空段落上（否则正文顶部会多一个空行）。
    """
    idx = _wd_para_index_at(doc, pos)
    if idx <= 0:
        return pos
    try:
        count = int(doc.Paragraphs.Count)
    except Exception:
        return pos
    for i in range(idx, count + 1):
        try:
            p = doc.Paragraphs(i)
            if _wd_para_text(p):
                return int(p.Range.Start)
        except Exception:
            break
    return pos


def _wd_page_anchor(doc, after_page: int) -> int:
    """返回「第 after_page 页之后」对应的字符位置（即第 after_page+1 页的页首）。

    页码按**插目录之前**的排版算 —— 界面上也是这么写的，用户看到的是他手上的稿子。
    """
    doc.Repaginate()
    try:
        total = int(doc.ComputeStatistics(2))      # wdStatisticPages
    except Exception:
        total = 0
    if total <= 0 or after_page >= total:
        return int(doc.Content.End) - 1            # 插在文档末尾之前
    try:
        rng = doc.GoTo(1, 1, int(after_page) + 1)  # wdGoToPage, wdGoToAbsolute
        return int(rng.Start)
    except Exception:
        return 0


def _wd_title_block_end(doc) -> int:
    """文档开头若是「标题 / 副标题」封面块，返回该块末尾的字符位置；否则返回 0。

    只作兜底：正常人会用 `_wd_toc_anchor()`，它靠第一个正文标题定位。
    这里留着是为了应付「封面用 Title 样式、正文却一个标题都没有」的文档。
    """
    try:
        count = int(doc.Paragraphs.Count)
    except Exception:
        return 0
    end = 0
    for i in range(1, min(count, 8) + 1):
        try:
            name = str(doc.Paragraphs(i).Style.NameLocal)
        except Exception:
            break
        if name in _COVER_STYLES:
            end = int(doc.Paragraphs(i).Range.End)
            continue
        break
    return end


def _wd_toc_anchor(doc, anchor=None) -> int:
    """算目录该插在哪个字符位置。

    ``anchor`` 三种取值：

    * ``None`` / ``"auto"``：自动 —— 优先「第一个正文标题之前」，兜底封面块末尾。
    * ``"start"``：文档最前面（封面之前）。
    * ``("page", N)``：第 N 页之后（按插目录前的排版算）。

    自动判断既然是「猜」，就必须留手动兜底：封面和正文挤在同一页、
    或者封面本身用了标题样式时，任何启发式都会出错，只有人知道该插哪。
    """
    if anchor == "start":
        return 0
    if isinstance(anchor, (tuple, list)) and len(anchor) == 2 and anchor[0] == "page":
        return _wd_page_anchor(doc, int(anchor[1]))

    pos = _wd_first_heading_pos(doc)
    if pos > 0:
        return pos
    return _wd_title_block_end(doc)


def _wd_is_toc_title_para(para) -> bool:
    """判断一个段落是不是目录标题行（「目录」/「目 录」/「Contents」）。"""
    try:
        text = str(para.Range.Text).replace("\r", "").replace("\x07", "")
    except Exception:
        return False
    flat = "".join(text.split())
    if flat in ("目录", "目録", "Contents", "CONTENTS", "TableofContents"):
        return True
    try:
        return "TOC" in str(para.Style.NameLocal).upper() and "标题" in str(para.Style.NameLocal)
    except Exception:
        return False


def _wd_toc_entry_count(toc) -> int:
    """数一个目录对象里有多少条目（空目录返回 0）。"""
    try:
        text = str(toc.Range.Text)
    except Exception:
        return 0
    if any(hint in text for hint in _TOC_EMPTY_HINTS):
        return 0
    return text.count("\r")


def _wd_ensure_outline_flag(toc) -> bool:
    """确保目录域里带 ``\\u``（按大纲级别收录）。改过域代码返回 True。

    为什么改域代码而不设属性：``TableOfContents.UseOutlineLevels`` 是**只读**的，
    赋值直接抛 AttributeError（实测），所以 ``toc.UseOutlineLevels = True``
    这种写法一直在 try/except 里静默失败 —— 等于没写。

    手动标记过的段落没有标题样式、只有大纲级别，域里不带 ``\\u`` 就一条都收不进去。
    Word 和 WPS 自己插的目录默认都带，这条是给「老目录 + 补标」兜底的。
    """
    try:
        field = toc.Range.Fields(1)
        code = str(field.Code.Text)
    except Exception:
        return False
    if "\\u" in code:
        return False
    try:
        field.Code.Text = code.rstrip() + " \\u "
        return True
    except Exception:
        return False


def _wd_update_existing_tocs(doc, levels: int, *, use_outline: bool = False) -> tuple[int, int]:
    """把已有目录**就地**改成 1..levels 级，返回 (更新了几个, 第一个的条目数)。

    只改几个属性 + Update()，不删不建：位置、缩进、字体、外面那层内容控件
    一个都不碰。实测 WPS 生成的目录，start 全程停在 36 没动过，条目 51 → 23。

    这是「文档本来就有目录」时的正确做法 —— 之前是删了重建，而重建的插入点
    算错就会把目录怼到封面前面去，等于把人家排好的版又搞乱了。

    use_outline=True 时额外让域代码带上「使用大纲级别」：手动标记过的段落没有
    标题样式，只有大纲级别，不打开这个选项目录里就看不到它们。
    """
    try:
        count = int(doc.TablesOfContents.Count)
    except Exception:
        return 0, 0

    done = 0
    entries = 0
    for i in range(1, count + 1):
        try:
            toc = doc.TablesOfContents(i)
            if use_outline:
                _wd_ensure_outline_flag(toc)   # 只读属性改不了，只能改域代码
            toc.UpperHeadingLevel = 1
            toc.LowerHeadingLevel = levels
            toc.Update()
            done += 1
            if i == 1:
                entries = _wd_toc_entry_count(toc)
        except Exception:
            continue
    return done, entries


def _wd_delete_tocs(doc, *, keep_first: bool = False) -> int:
    """删目录。keep_first=True 时只删第 2 个及以后的（重复生成的垃圾）。

    倒序删是为了不让前面的索引发生位移。
    """
    try:
        count = int(doc.TablesOfContents.Count)
    except Exception:
        return 0

    lowest = 2 if keep_first else 1
    removed = 0
    for i in range(count, lowest - 1, -1):
        try:
            toc = doc.TablesOfContents(i)
            start = int(toc.Range.Start)
            toc.Delete()
            removed += 1
            if start <= 0:
                continue
            prev = doc.Range(max(0, start - 1), start).Paragraphs(1)
            if str(prev.Range.Text).strip() == _TOC_TITLE:
                prev.Range.Delete()
        except Exception:
            continue
    return removed


def _wd_style_toc_heading(head, doc) -> None:
    """把「目录」那一行设成目录标题样式（居中、醒目、且不会被收进目录自己里）。

    为什么不用 Title / wdStyleTitle(-63)：Title 这个样式在不同模板里定义不一样，
    WPS 做的模板就把它带上了大纲级别，于是刚生成的目录第一行会多一条「目录……I」
    —— 目录把自己收进去了。实测踩过。

    「TOC 标题」（内置名 TOC Heading）是 Word 专为这件事准备的样式，
    既好看又保证不进目录；万一文档模板里没有这个样式，就退回手工排版。
    """
    for name in ("TOC Heading", "TOC 标题"):
        try:
            head.Style = doc.Styles(name)
            break
        except Exception:
            continue
    else:
        try:
            head.Range.Font.Bold = True
            head.Range.Font.Size = 16          # 三号
            head.Range.ParagraphFormat.Alignment = _WD_ALIGN_CENTER
        except Exception:
            pass

    # 双保险：就算模板给的样式带大纲级别，也按正文级别处理，不进目录
    try:
        head.Range.ParagraphFormat.OutlineLevel = 10   # wdOutlineLevelBodyText
    except Exception:
        pass


def _wd_insert_toc_into(doc, levels: int, *, with_title: bool = True,
                        page_break: bool = True,
                        replace_existing: bool = True,
                        use_outline: bool = False,
                        anchor=None) -> tuple[str, int, int]:
    """在已打开的 Word 文档里弄出一个原生目录，返回 (模式, 删掉/更新的旧目录数, 条目数)。

    两种模式，靠文档里**原来有没有目录**分流 —— 这是这件事的关键：

    * ``"updated"``：原来就有目录（Word、WPS 生成的都算），**原地改深度 + 刷页码**。
      位置、样式、外层内容控件全部不动。用户的排版是排过的，我们没有理由动它。
    * ``"inserted"``：原来没目录，才找位置新建。

    ``use_outline=True`` 时目录改成「以大纲级别为准」（域代码多一个 ``\\u``）。
    手动标记的段落没有标题样式、只有大纲级别，必须开这个开关才收得进去；
    对纯标题样式的文档没有副作用 —— 标题样式自带的大纲级别本来就在收录范围内。

    ``anchor`` 只在「新建」时有用，见 `_wd_toc_anchor()`：默认自动猜，
    也可以指定「文档最前面」或「第 N 页之后」。

    分页一律用段落属性 PageBreakBefore，**不插分页符字符**：
    删旧目录时带不走分页符，位置又会随域刷新漂移，重复生成会一次多一页
    （实测 4 页 → 5 页 → 6 页）。段落属性不占字符，天然幂等。

    更新只调 TablesOfContents(...).Update()，不调 doc.Fields.Update() ——
    目录里每条目都是一个超链接域，全文档刷域在大文档上明显更慢且没有收益。
    """
    levels = max(1, min(4, int(levels or 3)))

    # 空的大纲段落会被目录收成一条只有页码的怪条目，两头都要先清掉。
    # 放在最前面：这样「就地更新」那条路也受益（旧目录里的空条目会一起去掉）。
    _wd_flatten_empty_headings(doc)

    try:
        existing = int(doc.TablesOfContents.Count)
    except Exception:
        existing = 0

    # ---- 情况一：已经有目录 → 就地更新，位置一个字符都不动 ----
    if existing >= 1:
        if not replace_existing:
            # 用户明确说别碰已有目录 → 原样留着，什么都不做
            return "kept", 0, _wd_toc_entry_count(doc.TablesOfContents(1))
        if existing > 1:
            _wd_delete_tocs(doc, keep_first=True)      # 清掉重复生成的，只留第一个
        updated, entries = _wd_update_existing_tocs(doc, levels, use_outline=use_outline)
        doc.Repaginate()
        if entries > 0:
            return "updated", updated, entries
        # 原来的目录已经是个空壳（标题样式全丢了），只能删掉重建

    removed = _wd_delete_tocs(doc)

    # ---- 情况二：没有目录 → 找位置新建 ----
    pos = _wd_toc_anchor(doc, anchor)
    is_auto = anchor is None or anchor == "auto"
    if is_auto:
        pos = _wd_back_over_empty(doc, pos)   # 别把封面和正文之间的空段落顶成空页

    # 「目录」这一行要不要「段前分页」：
    #   · 自动 → 要。封面可能和正文挤在同一页，得让目录另起一页。
    #   · 最前面 → 不要，它就在文档第一页开头。
    #   · 指定页之后 → 不要，那个位置已经是某一页的第一个字符了，
    #     再设分页反而会把前面那页的空段落孤立成一张空白页（实测踩过）。
    need_break_before = is_auto and pos > 0

    if with_title:
        doc.Range(pos, pos).InsertBefore(_TOC_TITLE + "\r")
        head = doc.Range(pos, pos + len(_TOC_TITLE))
        _wd_style_toc_heading(head, doc)
        try:
            if need_break_before:
                head.ParagraphFormat.PageBreakBefore = _WD_ON   # 让封面单独一页
        except Exception:
            pass
        pos += len(_TOC_TITLE) + 1

    toc = doc.TablesOfContents.Add(
        doc.Range(pos, pos),   # Range：插入点
        True,                  # UseHeadingStyles：按标题样式建目录
        1,                     # UpperHeadingLevel
        levels,                # LowerHeadingLevel
        False,                 # UseFields：False = 用标题样式，不是 TC 域
        "",                    # TableID
        True,                  # RightAlignPageNumbers
        True,                  # IncludePageNumbers
        "",                    # AddedStyles
        True,                  # UseHyperlinks → 域里带 \h，Ctrl+点击可跳转
        False,                 # HidePageNumbersInWeb
        bool(use_outline),     # UseOutlineLevels：手动标记过的段落靠它收录
    )

    if page_break:
        try:
            end = int(toc.Range.End)
            if end < int(doc.Content.End):
                # 分页属性落在**正文第一段**上，而不是目录后面那个空段落上 ——
                # 落在空段上会让正文顶部多一个空行
                target = _wd_next_nonempty_pos(doc, end)
                if target < int(doc.Content.End):
                    doc.Range(target, target).ParagraphFormat.PageBreakBefore = _WD_ON
        except Exception:
            pass

    try:
        toc.Update()
    except Exception:
        try:
            doc.TablesOfContents(1).Update()
        except Exception:
            pass
    doc.Repaginate()

    return "inserted", removed, _wd_toc_entry_count(toc)


def add_toc_to_docx(src, out_path, *, levels: int = 3, with_title: bool = True,
                    page_break: bool = True, replace_existing: bool = True,
                    marks: dict | None = None, anchor=None, progress=None) -> dict:
    """复制一份 src 到 out_path，在**副本**里插目录 —— 源文件不参与任何写入。

    为什么用「复制文件 → 改副本」而不是「只读打开 → FormattedText 复制」：
    后者会丢页眉页脚和分节设置，而页码就住在页脚里，目录的页码依赖它。
    复制是最快的「零损失」办法（批注、修订、自定义属性也一并保住）。

    marks 是「手动标标题」的结果 {段落下标: 级别}：复制完先用 python-docx
    把大纲级别写进副本（只写 <w:outlineLvl>，不动样式和字体），再让 Word 打开。
    之所以不用 Word 的 Paragraphs(i) 去对号：python-docx 的 paragraphs 不含
    表格内段落，两边下标根本对不上，会标错行。

    anchor 决定目录插哪儿（只在文档原本没目录时生效）：
    ``None``/``"auto"`` 自动猜、``"start"`` 文档最前面、``("page", N)`` 第 N 页之后。

    返回 {"out": Path, "entries": int, "replaced": int, "mode": str, "anchor": str}
    """
    import pythoncom
    import win32com.client as win32

    src, out_path = Path(src), Path(out_path)
    if not src.is_file():
        raise RuntimeError(f"文件不存在：{src}")
    if src.suffix.lower() != ".docx":
        raise RuntimeError(f"只支持 .docx（这是 {src.suffix or '无扩展名'}），请先在 Word 里另存为 .docx。")
    if not has_word():
        raise RuntimeError("没检测到 Microsoft Word，算不出目录页码。")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink()
    if progress:
        progress(0, 4, "复制源文件…")
    shutil.copy2(src, out_path)          # 源文件一个字节都不动

    if marks:
        if progress:
            progress(1, 4, "写入标题级别…")
        from docx import Document

        from . import outline as outline_mod

        py_doc = Document(str(out_path))     # 这是副本，源文件仍然没被碰
        outline_mod.apply_levels_to_doc(py_doc, marks)
        py_doc.save(str(out_path))

    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = doc = None
        ok = False
        try:
            if progress:
                progress(2, 4, "启动 Word…")
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            try:
                app.Options.UpdateLinksAtOpen = False
            except Exception:
                pass
            # 12 个参数固定按位置：FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
            # PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument,
            # WritePasswordTemplate, Format, Encoding, Visible
            doc = app.Documents.Open(str(out_path), False, False, False, "",
                                     "", False, "", "", 0, 0, True)
            if progress:
                progress(3, 4, "生成目录并更新页码…")
            mode, removed, entries = _wd_insert_toc_into(
                doc, levels, with_title=with_title,
                page_break=page_break, replace_existing=replace_existing,
                use_outline=True,          # 总带 \u：纯样式文档无副作用，标记文档才收得到
                anchor=anchor,
            )
            if entries <= 0:
                raise RuntimeError(
                    "这份文档没有可识别的标题样式，Word 只会插出一个空目录。\n"
                    "请先把章节标题套上「标题 1 / 2 / 3」样式，或用「手动标标题」挑出章节，再试一次。"
                )
            doc.Save()
            ok = True
        finally:
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
            if not ok:
                try:
                    out_path.unlink()        # 失败别留下半成品
                except Exception:
                    pass

    if progress:
        if mode == "updated":
            progress(4, 4, f"完成：{out_path.name}（已就地更新原有目录，{entries} 条）")
        else:
            progress(4, 4, f"完成：{out_path.name}（新插入目录，{entries} 条）")
    return {"out": out_path, "entries": entries, "replaced": removed, "mode": mode}


def list_pages(src, *, limit_snippet: int = 40) -> list[dict]:
    """列出文档每一页的「页首文字」，给「目录插在第几页之后」用。

    为什么要列文字而不是只给页码：用户很难数清「封面是第 1 页还是第 2 页」，
    但一看「第 1 页 · 调查研究报告」就立刻知道该选哪个。

    只读打开，不写文件；没装 Word 返回空列表（页码必须靠排版算，
    python-docx 读不出分页位置，硬猜不如不给）。

    返回 ``[{"page": 1, "text": "调查研究报告"}, ...]``
    """
    import pythoncom
    import win32com.client as win32

    src = Path(src)
    if not src.is_file() or not has_word():
        return []

    pages: list[dict] = []
    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = doc = None
        try:
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            try:
                app.Options.UpdateLinksAtOpen = False
            except Exception:
                pass
            # 只读打开：第 3 个参数 ReadOnly=True
            doc = app.Documents.Open(str(src), False, True, False, "",
                                     "", False, "", "", 0, 0, True)
            doc.Repaginate()
            try:
                total = int(doc.ComputeStatistics(2))      # wdStatisticPages
            except Exception:
                total = 0
            if total <= 0:
                return []

            starts: list[int] = []
            for pg in range(1, total + 1):
                try:
                    starts.append(int(doc.GoTo(1, 1, pg).Start))   # wdGoToPage, wdGoToAbsolute
                except Exception:
                    starts.append(-1)
            end = int(doc.Content.End)

            for i, st in enumerate(starts):
                nxt = starts[i + 1] if i + 1 < len(starts) else end
                text = ""
                if st >= 0 and nxt > st:
                    try:
                        text = str(doc.Range(st, nxt).Text)
                    except Exception:
                        text = ""
                flat = " ".join(text.split())
                pages.append({"page": i + 1, "text": flat[:limit_snippet]})
        except Exception:
            return []
        finally:
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
    return pages


def add_toc_to_docx_pydocx(src, out_path, *, levels: int = 3, with_title: bool = True,
                           page_break: bool = True, marks: dict | None = None,
                           rows: list | None = None, progress=None) -> dict:
    """没有 Word 时的退路：直接写 TOC 域代码，并在 settings 里开 updateFields。

    **实测：Word 打开时不会自动更新这个域**，用户得手动 Ctrl+A → F9。
    所以界面必须如实说明，不能承诺「打开就有目录」。

    marks 给的是手动标记的 {段落下标: 级别}，会先写进文档再写域；
    rows 是已经排好序的 [(级别, 文本)]，用来填域结果（不传就自己扫一遍）。

    返回 {"out": Path, "entries": -1, "needs_update": True}
    """
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    src, out_path = Path(src), Path(out_path)
    if not src.is_file():
        raise RuntimeError(f"文件不存在：{src}")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    levels = max(1, min(4, int(levels or 3)))

    if progress:
        progress(0, 2, "写入目录域…")
    doc = Document(str(src))             # 只读 src，写的是 out_path
    if marks:
        from . import outline as outline_mod

        outline_mod.apply_levels_to_doc(doc, marks)
    first = doc.paragraphs[0] if doc.paragraphs else None

    def new_para():
        # 都插在同一个锚点前，后插的会紧贴锚点，所以顺序天然正确
        if first is not None:
            return first.insert_paragraph_before()
        return doc.add_paragraph()

    def run_with(para, el):
        run = para.add_run()
        run._r.append(el)
        return run

    if with_title:
        head = new_para()
        try:
            head.style = doc.styles["Title"]
        except Exception:
            pass
        head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        head.add_run(_TOC_TITLE)

    holder = new_para()

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")       # dirty 挂在 begin 上，某些版本会认
    run_with(holder, begin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")  # 不设的话前后空格会被吃掉，域代码就废了
    instr.text = f' TOC \\o "1-{levels}" \\h \\z \\u '
    run_with(holder, instr)

    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    run_with(holder, sep)

    # 域结果先填扫到的标题（用软换行，仍在一个段落里）：这样就算用户不按 F9，
    # 打开也能看到目录内容；按了 F9 就变成带点线和页码的正式目录。
    if rows is None:
        try:
            from . import toc as toc_mod

            rows = toc_mod.scan_headings(src)
        except Exception:
            rows = []
    rows = list(rows or [])[:80]
    if rows:
        result = holder.add_run()
        for i, (level, text) in enumerate(rows):
            if i:
                result.add_break()                       # 软换行
            result.add_text("　　" * (level - 1) + text)  # 全角空格缩进表现层级
    else:
        holder.add_run("（打开后按 Ctrl+A、F9 生成目录）")

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_with(holder, end)

    if page_break:
        holder.add_run().add_break(WD_BREAK.PAGE)

    # settings.xml 有固定的元素顺序要求，updateFields 必须落在 compat 之前
    try:
        settings_el = doc.settings.element
        update = OxmlElement("w:updateFields")
        update.set(qn("w:val"), "true")
        anchor = None
        for tag in ("w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr",
                    "w:compat", "w:docVars", "w:rsids"):
            found = settings_el.find(qn(tag))
            if found is not None:
                anchor = found
                break
        if anchor is not None:
            anchor.addprevious(update)
        else:
            settings_el.append(update)
    except Exception:
        pass

    doc.save(str(out_path))
    if progress:
        progress(2, 2, f"完成：{out_path.name}")
    return {"out": out_path, "entries": -1, "needs_update": True}


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


# --------------------------------------------------------------------------- PPT

# PowerPoint 的格式常量（win32com 动态派发下不能按名字传参，只能写数值）
_PP_SAVE_PDF = 32        # ppSaveAsPDF
_PP_SAVE_PPTX = 24       # ppSaveAsOpenXMLPresentation


def ppt_to_pdf(src: Path, *, force: bool = False) -> Path:
    """把演示文稿转成 PDF，一页幻灯片正好对应一页 PDF；按 (路径, 大小, 修改时间) 缓存。"""
    src = Path(src)
    st = src.stat()
    tag = _tag_of(str(src).lower(), st.st_size, int(st.st_mtime))
    out = cache_dir("ppt") / f"{src.stem}_{tag}.pdf"
    if out.exists() and out.stat().st_size > 0 and not force:
        return out
    if has_powerpoint():
        _ppt_export(src, out)
    elif soffice_path():
        _soffice_convert(src, out, "pdf")
    else:
        raise RuntimeError("未检测到 Microsoft PowerPoint 或 LibreOffice，无法解析 PPT 的幻灯片。")
    if not out.exists() or out.stat().st_size == 0:
        raise RuntimeError("PPT 转 PDF 失败，输出为空。")
    return out


def _ppt_export(src: Path, out: Path) -> None:
    """用 PowerPoint 把演示文稿另存为 PDF。

    踩过的坑：
      · Application.Visible 不能设成 False，PowerPoint 会直接抛
        "Hiding the application window is not allowed."；
        不想弹窗只能靠 Open 的 WithWindow 参数传 0（不开窗口）。
      · 动态派发下关键字参数会被静默丢掉，所以全部按位置传。
    """
    import pythoncom
    import win32com.client as win32

    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = None
        pres = None
        try:
            app = win32.DispatchEx("PowerPoint.Application")
            pres = app.Presentations.Open(str(src), -1, 0, 0)   # ReadOnly, Untitled, WithWindow
            pres.SaveAs(str(out), _PP_SAVE_PDF)
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


def ppt_extract_slides(src: Path, indexes: Iterable[int], out: Path) -> None:
    """把选中的幻灯片（**从 1 开始**）另存为一个新的 .pptx。

    实现上不走逐页 InsertFromFile —— 那样会把目标演示文稿的母版和配色套上去，
    导出的文件字体、版式全变样。这里改成「整份复制一份，再把不要的页删掉」，
    母版、版式、动画、备注都原样保留，代价只是文件里会留下一堆用不到的版式。
    """
    keep = sorted({int(i) for i in indexes})
    if not keep:
        raise ValueError("未选择任何幻灯片。")
    if not has_powerpoint():
        raise RuntimeError("导出 PowerPoint 需要安装 Microsoft PowerPoint。")
    _ppt_keep_slides(src, keep, out)
    if not out.exists() or out.stat().st_size == 0:
        raise RuntimeError("导出 PowerPoint 失败，输出为空。")


def _ppt_keep_slides(src: Path, keep: list[int], out: Path) -> None:
    import pythoncom
    import win32com.client as win32

    with _COM_LOCK:
        pythoncom.CoInitialize()
        app = None
        pres = None
        try:
            app = win32.DispatchEx("PowerPoint.Application")
            pres = app.Presentations.Open(str(src), -1, 0, 0)
            total = int(pres.Slides.Count)
            want = {i for i in keep if 1 <= i <= total}
            # 从后往前删，前面几页的序号才不会跟着变
            for i in range(total, 0, -1):
                if i not in want:
                    pres.Slides(i).Delete()
            if int(pres.Slides.Count) == 0:
                raise RuntimeError("所选幻灯片不存在。")
            pres.SaveAs(str(out), _PP_SAVE_PPTX)
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
