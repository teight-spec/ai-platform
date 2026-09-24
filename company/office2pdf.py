#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
office2pdf.py —— 用 LibreOffice 把 pptx/ppt/docx/doc/xlsx/xls 转成 PDF

用法:
    python3 /opt/company/office2pdf.py 文件 [更多文件...] [--keep] [--outdir 目录]

规则:
  * PDF 生成在原文件同目录（或 --outdir），同名 .pdf
  * PPT（.pptx/.ppt）转换成功后，原文件移到同目录「底稿」子文件夹——给用户的只有 PDF
    加 --keep 则原文件不动
  * 每次转换用独立的 LibreOffice 配置目录，多人同时转换不会互相卡住
  * 单个文件超过 180 秒算失败
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile

OK_EXT = {".pptx", ".ppt", ".docx", ".doc", ".xlsx", ".xls", ".odt", ".odp", ".ods", ".rtf"}
PPT_EXT = {".pptx", ".ppt"}


def convert(src, outdir, keep, timeout=180):
    src = os.path.abspath(src)
    base, ext = os.path.splitext(src)
    ext = ext.lower()
    if ext not in OK_EXT:
        return False, f"不支持的格式：{src}"
    if not os.path.isfile(src):
        return False, f"文件不存在：{src}"
    outdir = os.path.abspath(outdir or os.path.dirname(src))
    os.makedirs(outdir, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="lo_profile_", dir="/tmp")
    work = tempfile.mkdtemp(prefix="lo_out_", dir="/tmp")
    try:
        cmd = ["soffice", f"-env:UserInstallation=file://{profile}", "--headless",
               "--norestore", "--convert-to", "pdf", "--outdir", work, src]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired:
            return False, f"转换超时（>{timeout} 秒）：{src}"
        produced = os.path.join(work, os.path.basename(base) + ".pdf")
        if not os.path.isfile(produced) or os.path.getsize(produced) == 0:
            return False, f"转换失败：{src}\n{(r.stdout or '')[-500:]}{(r.stderr or '')[-500:]}"
        dst = os.path.join(outdir, os.path.basename(base) + ".pdf")
        shutil.move(produced, dst)
        msg = f"已生成：{dst}"
        if ext in PPT_EXT and not keep:
            draft_dir = os.path.join(os.path.dirname(src), "底稿")
            os.makedirs(draft_dir, exist_ok=True)
            moved = os.path.join(draft_dir, os.path.basename(src))
            shutil.move(src, moved)
            msg += f"\n  原 PPT 已移到：{moved}"
        return True, msg
    finally:
        shutil.rmtree(profile, ignore_errors=True)
        shutil.rmtree(work, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description="Office 文件转 PDF")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--keep", action="store_true", help="PPT 转换后原文件不移动")
    ap.add_argument("--outdir", help="PDF 输出目录（默认与原文件同目录）")
    a = ap.parse_args()
    if shutil.which("soffice") is None:
        print("终端缺少 LibreOffice（soffice），无法转换。")
        sys.exit(2)
    bad = 0
    for f in a.files:
        ok, msg = convert(f, a.outdir, a.keep)
        print(msg)
        bad += 0 if ok else 1
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
