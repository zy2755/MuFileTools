# -*- coding: utf-8 -*-
"""打印 Gradle 单元测试报告里的失败详情。"""
import re
import sys
from pathlib import Path

d = Path(r"D:\hou33\Documents\AutoClaw\软件\android\app\build\test-results\testDebugUnitTest")
for xml in d.glob("*.xml"):
    text = xml.read_text(encoding="utf-8", errors="replace")
    for m in re.finditer(r'<testcase name="([^"]+)"[^>]*>(.*?)</testcase>', text, re.S):
        name, body = m.group(1), m.group(2)
        fm = re.search(r"<failure[^>]*>(.*?)</failure>", body, re.S)
        if fm:
            msg = re.sub(r"\s+at [^\s>]+", "", fm.group(1)).strip()
            print("TEST:", name[:80])
            print("MSG :", msg[:500])
            print("---")
