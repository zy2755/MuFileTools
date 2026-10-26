"""在弱网环境下把需要的 wheel 拉到本地 wheels/ 目录（分片并行 + 续传 + 校验）。

用法： python tools/fetch_wheels.py [包名...]
随后： python -m pip install --no-index --find-links wheels pymupdf openpyxl pywin32 PySide6-Essentials
"""
from __future__ import annotations

import hashlib
import re
import sys
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import unquote, urljoin

MIRRORS = [
    "https://mirrors.aliyun.com/pypi/simple/{pkg}/",
    "https://pypi.tuna.tsinghua.edu.cn/simple/{pkg}/",
]
DEFAULT_PKGS = ["pymupdf", "openpyxl", "et-xmlfile", "pywin32", "PySide6-Essentials", "shiboken6"]
OUT = Path(__file__).resolve().parents[1] / "wheels"
WORKERS = 6
UA = {"User-Agent": "pip/24.0"}

_lock = threading.Lock()
_bytes = 0


def _get(url: str, timeout: int = 30):
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout)


def _version_key(v: str) -> tuple:
    return tuple(int(x) for x in re.findall(r"\d+", v))


def pick_wheel(page_html: str, base: str) -> tuple[str, str] | None:
    """从 simple 索引里挑一个适配 CPython 3.13 / win_amd64 的 wheel。"""
    best = None
    for href, text in re.findall(r'<a\s+href="([^"]+)"[^>]*>([^<]+)</a>', page_html):
        fn = text.strip()
        if not fn.endswith(".whl"):
            continue
        parts = fn[:-4].split("-")
        if len(parts) < 5:
            continue
        py, abi, plat = parts[-3], parts[-2], parts[-1]
        version = parts[1]
        if plat == "win_amd64":
            if "cp313" in py:
                score = 5
            elif "abi3" in abi and any(t in py for t in ("cp311", "cp312", "cp39", "cp310")):
                score = 4
            else:
                continue
        elif plat == "any":
            if not any(t in py for t in ("py3", "py2.py3")):
                continue
            score = 1
        else:
            continue
        sha = ""
        if "#sha256=" in href:
            sha = href.split("#sha256=")[1]
        # 优先正式版：预发布（a/b/rc/dev）永远排在正式版之后（key 越大越优，正式版取 1）
        pre = 0 if re.search(r"(a|b|rc|dev)\d?$", version.lower()) else 1
        key = (pre, _version_key(version), score)
        if best is None or key > best[0]:
            best = (key, urljoin(base, href.split("#")[0]), fn, sha)
    if not best:
        return None
    return best[1], best[2], best[3]


def probe_size(url: str) -> int:
    req = urllib.request.Request(url, headers={**UA, "Range": "bytes=0-0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        cr = r.headers.get("Content-Range")
        if cr and "/" in cr:
            return int(cr.split("/")[-1])
        return int(r.headers.get("Content-Length") or 0)


def _fetch_part(url: str, dest: Path, start: int, end: int) -> None:
    global _bytes
    pos = start
    tries = 40
    while pos <= end:
        try:
            req = urllib.request.Request(url, headers={**UA, "Range": f"bytes={pos}-{end}"})
            with urllib.request.urlopen(req, timeout=45) as r, open(dest, "r+b") as f:
                f.seek(pos)
                while pos <= end:
                    chunk = r.read(131072)
                    if not chunk:
                        break
                    f.write(chunk)
                    pos += len(chunk)
                    with _lock:
                        _bytes += len(chunk)
        except Exception:
            if tries <= 0:
                raise
            tries -= 1
            time.sleep(0.8)


def download(url: str, dest: Path, sha: str = "") -> None:
    global _bytes
    size = probe_size(url)
    if size <= 0:
        raise RuntimeError(f"无法获取文件大小：{url}")
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as f:
        f.truncate(size)

    step = max(size // WORKERS, 512 * 1024)
    spans = [(s, min(s + step - 1, size - 1)) for s in range(0, size, step)]
    _bytes = 0
    t0 = time.time()
    stop = threading.Event()

    def ticker():
        while not stop.wait(10):
            print(f"    … {_bytes / 1048576:6.2f}/{size / 1048576:.1f} MB "
                  f"({_bytes / 1048576 / max(time.time() - t0, .1):.2f} MB/s)", flush=True)

    th = threading.Thread(target=ticker, daemon=True)
    th.start()
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futs = [pool.submit(_fetch_part, url, dest, a, b) for a, b in spans]
        for f in futs:
            f.result()
    stop.set()

    got = dest.stat().st_size
    if got != size:
        raise RuntimeError(f"大小不符：{got} != {size}")
    if sha:
        h = hashlib.sha256()
        with open(dest, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        if h.hexdigest() != sha:
            raise RuntimeError("sha256 校验失败")
    print(f"    ok {dest.name}  {size / 1048576:.1f} MB  "
          f"{size / 1048576 / max(time.time() - t0, .1):.2f} MB/s", flush=True)


def resolve(pkg: str) -> tuple[str, str, str]:
    last: Exception | None = None
    for tpl in MIRRORS:
        idx = tpl.format(pkg=pkg.lower())
        try:
            with _get(idx) as r:
                page = r.read().decode("utf-8", "ignore")
            picked = pick_wheel(page, idx)
            if picked:
                return picked
        except Exception as exc:  # noqa: BLE001
            last = exc
    raise RuntimeError(f"在镜像里没找到 {pkg} 的可用 wheel：{last}")


def main() -> int:
    pkgs = sys.argv[1:] or DEFAULT_PKGS
    OUT.mkdir(parents=True, exist_ok=True)
    for pkg in pkgs:
        try:
            url, fn, sha = resolve(pkg)
        except Exception as exc:
            print(f"[{pkg}] 解析失败：{exc}", flush=True)
            continue
        dest = OUT / unquote(fn)
        if dest.exists() and dest.stat().st_size > 0:
            print(f"[{pkg}] 已存在 {dest.name}", flush=True)
            continue
        print(f"[{pkg}] {fn}", flush=True)
        for attempt in range(1, 4):
            try:
                download(url, dest, sha)
                break
            except Exception as exc:
                print(f"    第 {attempt} 次失败：{exc}", flush=True)
                dest.unlink(missing_ok=True)
                if attempt == 3:
                    print(f"[{pkg}] 放弃", flush=True)
    print("DONE", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
