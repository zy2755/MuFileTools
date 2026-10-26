"""导出器：把选中的单位写成图片 / PDF / Word / Excel。

对外只有 run_jobs(jobs, spec, progress) -> [输出文件]
"""
from __future__ import annotations

import io
import re
from pathlib import Path

from PIL import Image

from .models import ExportSpec, Job

_BAD_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]+')
MAX_UNIT_TAG = 60


# --------------------------------------------------------------------------- 入口

def run_jobs(jobs: list[Job], spec: ExportSpec, progress=None) -> list[Path]:
    """按 spec 导出。progress(done, total, message) 用于上报进度。"""
    jobs = [j for j in jobs if j.units]
    if not jobs:
        raise ValueError("没有选中任何页面或工作表。")

    out_dir = Path(spec.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    fmt = spec.fmt.lower()

    total = 0
    for job in jobs:
        total += len(job.units) if fmt in ("png", "jpg") and not spec.long_image else len(job.units)
    total = max(total, 1)

    results: list[Path] = []
    done = 0
    skipped: list[str] = []

    for job in jobs:
        src = job.source
        kind = getattr(src, "kind", "")
        stem = _clean(Path(src.path).stem)

        if fmt in ("png", "jpg"):
            if spec.long_image:
                imgs: list[Image.Image] = []
                for u in job.units:
                    imgs.extend(src.render(u.key, spec.dpi))
                    done += 1
                    _tick(progress, done, total, f"渲染 {stem} · {u.label}")
                target = _unique(out_dir / f"{spec.prefix}{stem}_长图{spec.suffix}.{fmt}", spec.overwrite)
                if target is None:
                    skipped.append(stem)
                    continue
                _stitch(imgs).save(target, **_save_kwargs(fmt, spec))
                results.append(target)
            else:
                for u in job.units:
                    pages = src.render(u.key, spec.dpi)
                    for i, img in enumerate(pages, start=1):
                        suffix = "" if len(pages) == 1 else f"_p{i}"
                        name = f"{spec.prefix}{stem}_{_clean(u.label) if kind == 'excel' else u.label.replace(' 页', '')}{suffix}{spec.suffix}.{fmt}"
                        target = _unique(out_dir / _clean(name.replace(" ", "")), spec.overwrite)
                        if target is None:
                            skipped.append(name)
                            continue
                        img.save(target, **_save_kwargs(fmt, spec))
                        results.append(target)
                    done += 1
                    _tick(progress, done, total, f"导出 {stem} · {u.label}")

        elif fmt == "pdf":
            target = _unique(out_dir / f"{spec.prefix}{stem}_提取{spec.suffix}.pdf", spec.overwrite)
            if target is None:
                skipped.append(stem)
                done += len(job.units)
                continue
            src.export_pdf(job.keys, target)
            results.append(target)
            done += len(job.units)
            _tick(progress, done, total, f"写出 {target.name}")

        elif fmt == "docx":
            target = _unique(out_dir / f"{spec.prefix}{stem}_提取{spec.suffix}.docx", spec.overwrite)
            if target is None:
                skipped.append(stem)
                done += len(job.units)
                continue
            blocks: list[list[Image.Image]] = []
            for u in job.units:
                blocks.append(src.render(u.key, spec.dpi))
                done += 1
                _tick(progress, done, total, f"排版 {stem} · {u.label}")
            _build_docx(blocks, target, title=f"{stem} 提取")
            results.append(target)

        elif fmt == "xlsx":
            if kind != "excel":
                done += len(job.units)
                _tick(progress, done, total, f"跳过 {stem}（仅 Excel 支持 xlsx 导出）")
                continue
            if spec.merge_xlsx:
                target = _unique(out_dir / f"{spec.prefix}{stem}_工作表{spec.suffix}.xlsx", spec.overwrite)
                if target is None:
                    skipped.append(stem)
                    done += len(job.units)
                    continue
                src.export_workbook(job.keys, target)
                results.append(target)
                done += len(job.units)
                _tick(progress, done, total, f"写出 {target.name}")
            else:
                for u in job.units:
                    target = _unique(out_dir / f"{spec.prefix}{stem}_{_clean(u.label)}{spec.suffix}.xlsx", spec.overwrite)
                    done += 1
                    if target is None:
                        skipped.append(u.label)
                        continue
                    src.export_workbook([u.key], target)
                    results.append(target)
                    _tick(progress, done, total, f"写出 {target.name}")
        else:
            raise ValueError(f"未知导出格式：{fmt}")

    if progress:
        msg = f"完成，共生成 {len(results)} 个文件"
        if skipped:
            msg += f"；跳过 {len(skipped)} 个同名文件"
        progress(total, total, msg)
    return results


# --------------------------------------------------------------------------- 辅助

def _tick(progress, done: int, total: int, msg: str) -> None:
    if progress:
        progress(min(done, total), total, msg)


def _save_kwargs(fmt: str, spec: ExportSpec) -> dict:
    if fmt == "jpg":
        return {"quality": max(50, min(100, spec.jpg_quality)), "optimize": True, "subsampling": 0}
    return {"optimize": True}


def _clean(name: str) -> str:
    out = _BAD_CHARS.sub("_", name).strip(" .")
    return (out or "未命名")[:MAX_UNIT_TAG]


def _unique(path: Path, mode: str) -> Path | None:
    """mode: overwrite 直接覆盖 / skip 返回 None / rename 自动加序号。"""
    if not path.exists():
        return path
    if mode == "overwrite":
        return path
    if mode == "skip":
        return None
    i = 2
    while True:
        cand = path.with_name(f"{path.stem}({i}){path.suffix}")
        if not cand.exists():
            return cand
        i += 1


def _stitch(images: list[Image.Image], gap: int = 14, bg: str = "white") -> Image.Image:
    """纵向拼接长图，宽度不足处补白边。"""
    if not images:
        raise ValueError("没有可拼接的图片。")
    imgs = [im.convert("RGB") for im in images]
    width = max(im.width for im in imgs)
    height = sum(im.height for im in imgs) + gap * (len(imgs) - 1)
    canvas = Image.new("RGB", (width, height), bg)
    y = 0
    for im in imgs:
        canvas.paste(im, ((width - im.width) // 2, y))
        y += im.height + gap
    return canvas


def _build_docx(blocks: list[list[Image.Image]], out_path: Path, title: str = "") -> None:
    from docx import Document
    from docx.enum.text import WD_BREAK
    from docx.shared import Cm

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.2)
    sec.top_margin = sec.bottom_margin = Cm(1.2)
    usable = sec.page_width - sec.left_margin - sec.right_margin

    first = True
    for pages in blocks:
        for img in pages:
            if not first:
                # 每个单位（原始一页/一张表）另起一页，避免内容错位
                doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            first = False
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format="PNG", optimize=True)
            buf.seek(0)
            pic = doc.add_picture(buf, width=usable)
            pic.alignment = 1  # 居中
    doc.save(str(out_path))
