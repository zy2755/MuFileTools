"""核心数据模型：可选单位（页 / 工作表）、导出参数、任务。"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Literal

Kind = Literal["pdf", "word", "excel", "ppt"]

KIND_LABEL: dict[str, str] = {"pdf": "PDF", "word": "Word", "excel": "Excel", "ppt": "PPT"}

SUPPORTED_EXTS: dict[str, Kind] = {
    ".pdf": "pdf",
    ".doc": "word",
    ".docx": "word",
    ".docm": "word",
    ".rtf": "word",
    ".xls": "excel",
    ".xlsx": "excel",
    ".xlsm": "excel",
    ".xlsb": "excel",
    ".ppt": "ppt",
    ".pptx": "ppt",
    ".pptm": "ppt",
    ".pps": "ppt",
    ".ppsx": "ppt",
}


@dataclass(frozen=True)
class Unit:
    """一个可勾选的单位：PDF/Word/PPT 的一页，或 Excel 的一张工作表。"""

    key: str
    label: str
    meta: dict[str, Any] = field(default_factory=dict)


@dataclass
class ExportSpec:
    fmt: str = "png"                 # png | jpg | pdf | docx | xlsx | pptx
    out_dir: Path = field(default_factory=lambda: Path.cwd() / "导出")
    dpi: int = 200
    jpg_quality: int = 92
    long_image: bool = False         # 仅 png/jpg：把所选单位拼成一张长图
    merge_xlsx: bool = True          # 仅 xlsx：多张工作表合并进一个工作簿
    merge_docx: bool = True          # 仅 docx：跨文件合并成一份文档（且允许调顺序）
    overwrite: str = "rename"        # rename | skip | overwrite
    prefix: str = ""
    suffix: str = ""


@dataclass
class Job:
    """一次导出的最小任务：某个文件里被选中的若干单位。"""

    source: Any
    units: list[Unit]

    @property
    def keys(self) -> list[str]:
        return [u.key for u in self.units]


@dataclass
class PageRef:
    """合并导出用的一页：来自哪个文件、第几页。列表顺序 = 输出顺序。"""

    source: Any
    key: str
    label: str
    doc_name: str

    @property
    def kind(self) -> str:
        return getattr(self.source, "kind", "")


ProgressFn = Callable[[int, int, str], None]
