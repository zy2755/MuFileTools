# -*- coding: utf-8 -*-
"""手动标记标题级别 —— 给「一个标题样式都没套」的文档补一条后路。

WPS 用户的重灾区：文档明明有「第一章」「一、」「1.1」这些标题，
但全都是手打的（正文样式 + 加粗 + 大字），Word 的目录一条都认不出来。

这里的做法是**只写一个 <w:outlineLvl>**（段落属性里的大纲级别），
不碰样式、不碰字体、不碰字号颜色、不碰缩进 —— 屏幕上什么都不会变，
但 Word 的目录和导航窗格都认它。

为什么不是「套用标题 1/2/3 样式」：那样会把用户排好的字体、字号、行距、
颜色全换成 Word 的内置样式，一份调了半天的版直接面目全非。
大纲级别是藏在段落属性里的元数据，属于「说了但不改」的那种修改。

只做三件事：
  list_paragraphs()  读 —— 列出所有非空段落，带字号/加粗/对齐/现有级别
  guess_levels()     猜 —— 按编号规律 + 字形特征预判一遍，人再微调
  apply_levels()     写 —— 把标记写进副本，源文件一个字节不动
"""
from __future__ import annotations

import re
import statistics
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from . import toc

MAX_LEVEL = 4
DEFAULT_SIZE = 10.5          # 读不到字号时的兜底（五号）
_MAX_GUESS_LEN = 60          # 超过这个长度的一律不猜（太像正文）


@dataclass(frozen=True)
class Para:
    """列表里的一行。index 是 doc.paragraphs 的下标，回写时靠它定位。"""

    index: int
    text: str
    style: str
    size: float
    bold: bool
    centered: bool
    level: int          # 文档里现有的大纲级别（0 = 正文）
    candidate: bool     # 像不像标题：给界面的「只看候选」筛选用


# --------------------------------------------------------------------------- 读

def _run_size(p) -> float | None:
    """取段落里第一个显式设了字号的 run；都继承样式则返回 None。"""
    for r in p.runs:
        size = r.font.size
        if size is not None:
            return float(size.pt)
    return None


def _style_size(style) -> float | None:
    """沿继承链找字号。python-docx 的 style.font.size 没显式设置时是 None。"""
    depth = 0
    while style is not None and depth < 8:
        try:
            if style.font.size is not None:
                return float(style.font.size.pt)
        except Exception:
            return None
        try:
            style = style.base_style
        except Exception:
            return None
        depth += 1
    return None


def _run_bold(p) -> bool | None:
    for r in p.runs:
        if r.font.bold is not None:
            return bool(r.font.bold)
    return None


def _style_bold(style) -> bool | None:
    depth = 0
    while style is not None and depth < 8:
        try:
            if style.font.bold is not None:
                return bool(style.font.bold)
        except Exception:
            return None
        try:
            style = style.base_style
        except Exception:
            return None
        depth += 1
    return None


def _font_of(p) -> tuple[float, bool]:
    """（字号 pt, 是否加粗）。run 直接格式优先，其次沿样式链继承。"""
    try:
        style = p.style
    except Exception:
        style = None
    size = _run_size(p)
    if size is None:
        size = _style_size(style)
    if size is None:
        size = DEFAULT_SIZE

    bold = _run_bold(p)
    if bold is None:
        bold = _style_bold(style)
    return float(size), bool(bold)


def list_paragraphs(path: str | Path) -> list[Para]:
    """列出所有非空段落（空段落不列 —— 它不可能成为标题）。

    已知边界：doc.paragraphs 不含表格单元格里的段落，所以表格里的标题
    在列表里看不到（Word 插目录时其实会收录）。这里宁少不多。

    表格是重灾区里的重灾区（很多报告用表格排版封面），所以额外给
    表格里的段落也编了号 —— 用负索引区分，回写时走同一条路径。
    """
    path = Path(path)
    try:
        doc = Document(str(path))
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"读不了这份文档：{exc}") from exc

    raw: list[tuple[int, object, str]] = []
    for i, p in enumerate(doc.paragraphs):
        raw.append((i, p, "body"))
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                for pi, p in enumerate(cell.paragraphs):
                    # 负索引编码：-1 - (表号*10000 + 行*100 + 列*10 + 段号)
                    raw.append((-(1 + ti * 10000 + ri * 100 + ci * 10 + pi), p, "table"))

    sized: list[tuple[Para, float]] = []
    sizes: list[float] = []
    for idx, p, _kind in raw:
        text = (p.text or "").strip()
        if not text:
            continue
        try:
            style_name = str(p.style.name or "")
        except Exception:
            style_name = ""
        # 目录条目本身也是段落（Word 原生目录不在内容控件里），必须剔掉，
        # 否则用户会在清单里看到「1 项目简介.........2」这种行。
        if toc.is_toc_style(style_name):
            continue
        size, bold = _font_of(p)
        sizes.append(size)
        try:
            centered = p.alignment == WD_ALIGN_PARAGRAPH.CENTER
        except Exception:
            centered = False
        level = toc.level_of(p) or 0
        if level > MAX_LEVEL:
            level = MAX_LEVEL
        sized.append((
            Para(index=idx, text=text, style=style_name, size=size, bold=bold,
                 centered=centered, level=level, candidate=False),
            size,
        ))

    body = float(statistics.median(sizes)) if sizes else DEFAULT_SIZE

    out: list[Para] = []
    for para, _size in sized:
        cand = _looks_like_heading(para, body)
        out.append(Para(
            index=para.index, text=para.text, style=para.style, size=para.size,
            bold=para.bold, centered=para.centered, level=para.level, candidate=cand,
        ))
    return out


# --------------------------------------------------------------------------- 猜

# 编号规律：从强到弱依次匹配，命中最强的那个
_RE_CHAPTER = re.compile(r"^第\s*[一二三四五六七八九十百零〇\d]+\s*[章节篇部讲]")
_RE_SPECIAL = re.compile(r"^(附\s*录|参考文献|致谢|结语|结论|内容摘要|摘\s*要|前\s*言|引\s*言)")
_RE_ARABIC = re.compile(r"^(\d+(?:[.．]\d+)*)[\s、．.]+(?=\S)")
_RE_CN_PAREN = re.compile(r"^[（(]\s*[一二三四五六七八九十]+\s*[)）]")
_RE_CN_NUM = re.compile(r"^[一二三四五六七八九十]+\s*[、．.]\s*")
# 纯日期 / 纯编号的行（封面的「2026年2月12日」「1」）长得像大字标题，
# 但收进目录毫无意义。整段只由数字、年月日、符号和空格组成就判掉。
_RE_DATE_ONLY = re.compile(r"^[\d\s年月日\-—–~～./、,，.。:：]+$")


def _looks_like_heading(p: Para, body: float) -> bool:
    """像不像标题 —— 只用来给界面「只看候选」过滤，不影响最终判定。"""
    text = p.text
    if len(text) > _MAX_GUESS_LEN:
        return False
    if _RE_DATE_ONLY.match(text):
        return False
    if p.level > 0:
        return True
    if _guess_text(text) > 0:
        return True
    if p.size - body >= 2 and len(text) <= 30:
        return True
    if p.bold and len(text) <= 20:
        return True
    return False


def _guess_text(text: str) -> int:
    """纯按文字规律猜级别，猜不到返回 0。"""
    if not text:
        return 0
    if _RE_CHAPTER.match(text) or _RE_SPECIAL.match(text):
        return 1
    m = _RE_ARABIC.match(text)
    if m:
        dots = m.group(1).count(".") + m.group(1).count("．")
        return min(dots + 1, MAX_LEVEL)
    if _RE_CN_PAREN.match(text):
        return 3
    if _RE_CN_NUM.match(text):
        return 2
    return 0


def guess_levels(paras: list[Para]) -> dict[int, int]:
    """自动预判，返回 {段落下标: 级别}，只含级别 > 0 的行。

    两条线并行，谁先命中听谁的：
      1. 编号规律（「第三章」「一、」「1.2.3」「（四）」）—— 最可靠
      2. 字形特征（比正文大、加粗、居中）—— 兜住手打的「第一点：」这种

    故意保守：宁可漏猜让人手动补，也不要把一句普通的正文标成标题 ——
    后者要在几十行里翻出来改掉，更烦。
    """
    if not paras:
        return {}
    body = float(statistics.median([p.size for p in paras])) if paras else DEFAULT_SIZE

    marks: dict[int, int] = {}
    for p in paras:
        text = p.text.strip()
        if len(text) > _MAX_GUESS_LEN:
            continue
        if _RE_DATE_ONLY.match(text):
            continue          # 封面的日期、编号，收进目录没意义
        if text and text[-1] in "。，；,;：:":
            continue          # 以标点收尾的十有八九是正文，不是标题

        level = _guess_text(text)
        if level == 0:
            big = p.size - body
            if p.centered and big >= 4:
                level = 1
            elif big >= 6:
                level = 1
            elif big >= 2 and (p.bold or p.centered):
                level = 2
        if level > 0:
            marks[p.index] = min(level, MAX_LEVEL)
    return marks


# --------------------------------------------------------------------------- 写

def _set_outline(p_el, level: int) -> bool:
    """给 <w:p> 写/删 <w:outlineLvl>。level 0 = 删掉（跟随样式）。返回是否改动。"""
    ppr = p_el.find(qn("w:pPr"))
    node = None if ppr is None else ppr.find(qn("w:outlineLvl"))

    if level <= 0:
        if node is not None:
            ppr.remove(node)
            return True
        return False

    if node is None:
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            p_el.insert(0, ppr)        # pPr 必须是 p 的第一个子元素
        node = OxmlElement("w:outlineLvl")
        # CT_PPrBase 是固定 sequence：outlineLvl 要排在这些之前
        anchor = None
        for tag in ("w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange"):
            found = ppr.find(qn(tag))
            if found is not None:
                anchor = found
                break
        if anchor is not None:
            anchor.addprevious(node)
        else:
            ppr.append(node)

    val = str(level - 1)               # 0..3 对应大纲 1..4 级
    if node.get(qn("w:val")) == val:
        return False
    node.set(qn("w:val"), val)
    return True


def apply_levels_to_doc(doc, marks: dict[int, int]) -> int:
    """把 {段落下标: 级别} 写进已打开的 python-docx 文档，返回改动条数。"""
    if not marks:
        return 0
    paras = list(doc.paragraphs)
    cells: list = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cells.extend(cell.paragraphs)

    changed = 0
    for idx, level in marks.items():
        if idx >= 0:
            if idx < len(paras):
                p = paras[idx]
            else:
                continue
        else:
            pos = (-idx) - 1
            ti, rest = divmod(pos, 10000)
            ri, rest = divmod(rest, 100)
            ci, pi = divmod(rest, 10)
            try:
                p = doc.tables[ti].rows[ri].cells[ci].paragraphs[pi]
            except Exception:
                continue
        try:
            if _set_outline(p._p, int(level)):
                changed += 1
        except Exception:
            continue
    return changed


def apply_levels(src, out_path, marks: dict[int, int]) -> int:
    """复制 src 到 out_path，把标记写进**副本**；返回改动条数。

    源文件全程不打开写入 —— 和 core.office 的目录生成保持同一套约定。
    """
    import shutil

    src, out_path = Path(src), Path(out_path)
    if not src.is_file():
        raise RuntimeError(f"文件不存在：{src}")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, out_path)

    doc = Document(str(out_path))
    changed = apply_levels_to_doc(doc, marks)
    doc.save(str(out_path))
    return changed
