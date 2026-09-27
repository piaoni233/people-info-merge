#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""多表人员信息汇总助手 · 打包脚本。

把技能目录打包为上传用 ZIP，并做上架前校验：
  1. ZIP 根层级为 `多表人员信息汇总助手/SKILL.md`（平台要求）。
  2. 总大小 ≤ 10MB。
  3. 排除运行产物与真实数据目录（profile.json/keys.json/plan.json/输出 xlsx/data/ 等）。
  4. 包内无 .env/*.key/*.pem（密钥兜底检查）。

用法（在技能目录的父目录或任意目录执行）：
  python scripts/package_skill.py [--out out.zip]
"""
from __future__ import annotations

import os
import sys
import zipfile

try:  # Windows 控制台中文输出保护（与 people_merge.py 保持一致）
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SKILL_DIR_NAME = "多表人员信息汇总助手"
MAX_BYTES = 10 * 1024 * 1024
EXCLUDE_DIRS = {"__pycache__", ".git", "data", "input", "output", ".venv", "venv"}
EXCLUDE_FILES = {"profile.json", "keys.json", "plan.json", "recipe.json",
                 "汇总结果.xlsx", "汇总结果.report.json", ".env"}
EXCLUDE_SUFFIX = (".zip", ".pyc", ".pyo", ".report.json")
FORBIDDEN_SUFFIX = (".key", ".pem", ".p12")


def skill_root() -> str:
    here = os.path.abspath(os.path.dirname(__file__))
    root = os.path.abspath(os.path.join(here, os.pardir))
    if os.path.basename(root) != SKILL_DIR_NAME:
        print("错误：请把本脚本放在 %s/scripts/ 下再运行。" % SKILL_DIR_NAME, file=sys.stderr)
        sys.exit(2)
    return root


def iter_files(root: str):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if d not in EXCLUDE_DIRS)
        for fn in sorted(filenames):
            if fn in EXCLUDE_FILES or fn.endswith(EXCLUDE_SUFFIX):
                continue
            if fn.endswith(FORBIDDEN_SUFFIX) or fn == ".env":
                print("错误：包内禁止包含密钥类文件：%s" % os.path.join(dirpath, fn), file=sys.stderr)
                sys.exit(2)
            if fn.endswith(".xlsx") and "examples" not in os.path.relpath(dirpath, root).split(os.sep):
                continue  # 仅排除运行产物 xlsx；examples 内如有示例 xlsx 可保留
            yield os.path.join(dirpath, fn)


def main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=SKILL_DIR_NAME + ".zip")
    a = ap.parse_args(argv)
    root = skill_root()
    skill_md = os.path.join(root, "SKILL.md")
    if not os.path.isfile(skill_md) or os.path.getsize(skill_md) == 0:
        print("错误：SKILL.md 缺失或为空，拒绝打包。", file=sys.stderr)
        return 2
    out = os.path.abspath(a.out)
    names = []
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for fp in iter_files(root):
            arc = os.path.join(SKILL_DIR_NAME, os.path.relpath(fp, root)).replace(os.sep, "/")
            zf.write(fp, arc)
            names.append(arc)
    size = os.path.getsize(out)
    print("打包完成：%s" % out)
    print("文件数：%d，总大小：%d 字节（%.2f MB，上限 10MB）"
          % (len(names), size, size / 1048576.0))
    assert SKILL_DIR_NAME + "/SKILL.md" in names, "ZIP 根层级缺少 SKILL.md！"
    print("校验通过：%s/SKILL.md 位于 ZIP 根层级。" % SKILL_DIR_NAME)
    if size > MAX_BYTES:
        print("错误：超过 10MB 上限，请压缩 assets 后重试。", file=sys.stderr)
        return 2
    for n in names:
        print("  " + n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
