#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
stamp_internal.py —— 给交付文件统一打上「内部文件，禁止外传」标注。

设计原则：只加标注，绝不改动原有数据结构。
  * 不插入行/列（openpyxl 插行不会自动修正公式，会毁掉可追溯性）
  * 不改动 sheet 顺序、不改表名
  * 幂等：重复运行不会叠加多次标注

用法:
    python3 stamp_internal.py 文件或目录 [更多...] [--dry-run] [--text "自定义文字"]
    python3 stamp_internal.py 输出目录/ -r        # 递归整个目录

支持: .xlsx .xlsm  .docx  .pptx  .html/.htm  .pdf
"""
import argparse
import os
import re
import sys

MARK = "内部文件，禁止外传"
GRAY = "808080"


# ---------------------------------------------------------------- xlsx
def stamp_xlsx(path, mark, dry=False):
    import openpyxl
    from openpyxl.styles import Font, Alignment

    wb = openpyxl.load_workbook(path)  # 保留公式（不加 data_only）
    touched = []

    for ws in wb.worksheets:
        # 1) 页脚：打印/页面布局视图可见，不占用任何单元格
        foot = ws.oddFooter.right.text or ""
        if mark not in foot:
            ws.oddFooter.right.text = (foot + "  " if foot else "") + mark
            ws.oddFooter.right.size = 9
            ws.oddFooter.right.color = GRAY
        ws.evenFooter.right.text = ws.oddFooter.right.text
        ws.firstFooter.right.text = ws.oddFooter.right.text

        # 2) 标题行：仅当 A1 是纯文本（非公式、非空）时，追加到标题末尾
        a1 = ws["A1"]
        v = a1.value
        if isinstance(v, str) and v.strip() and not v.startswith("="):
            if mark not in v:
                a1.value = f"{v}　【{mark}】"
                touched.append(ws.title)
        elif v is None:
            # A1 空 → 直接写标注（不影响任何公式）
            a1.value = f"【{mark}】"
            a1.font = Font(size=9, color=GRAY, bold=True)
            a1.alignment = Alignment(horizontal="left", vertical="center")
            touched.append(ws.title)
        # A1 是公式或数字 → 只靠页脚，不动它

    # 3) 文档属性
    try:
        wb.properties.keywords = mark
        wb.properties.category = mark
    except Exception:
        pass

    if not dry:
        wb.save(path)
    return f"xlsx: {len(wb.worksheets)} 页设页脚；标题行标注 {len(touched)} 页"


# ---------------------------------------------------------------- docx
def stamp_docx(path, mark, dry=False):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document(path)
    n = 0
    for section in doc.sections:
        for footer in (section.footer, section.first_page_footer, section.even_page_footer):
            if footer is None:
                continue
            if any(mark in p.text for p in footer.paragraphs):
                continue
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            if p.text.strip():
                p = footer.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(mark)
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
            n += 1

    # 正文首行也放一条，确保屏幕阅读时第一眼能看到
    body_has = any(mark in p.text for p in doc.paragraphs[:5])
    if not body_has and doc.paragraphs:
        first = doc.paragraphs[0]
        newp = first.insert_paragraph_before(mark)
        newp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = newp.runs[0]
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    try:
        doc.core_properties.keywords = mark
        doc.core_properties.category = mark
    except Exception:
        pass

    if not dry:
        doc.save(path)
    return f"docx: {n} 处页脚 + 正文首行标注"


# ---------------------------------------------------------------- pptx
def stamp_pptx(path, mark, dry=False):
    from pptx import Presentation
    from pptx.util import Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    prs = Presentation(path)
    W, H = prs.slide_width, prs.slide_height
    box_w = Emu(int(W * 0.30))
    box_h = Pt(16)
    left = Emu(int(W - box_w - Emu(int(W * 0.03))))
    top = Emu(int(H - box_h - Emu(int(H * 0.02))))

    n = 0
    for slide in prs.slides:
        already = False
        for shp in slide.shapes:
            if shp.has_text_frame and mark in shp.text_frame.text:
                already = True
                break
        if already:
            continue
        tb = slide.shapes.add_textbox(left, top, box_w, box_h)
        tf = tb.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        r = p.add_run()
        r.text = mark
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        n += 1

    try:
        prs.core_properties.keywords = mark
        prs.core_properties.category = mark
    except Exception:
        pass

    if not dry:
        prs.save(path)
    return f"pptx: {n} 页加标注（共 {len(prs.slides)} 页）"


# ---------------------------------------------------------------- html
HTML_BLOCK = """
<!-- internal-mark -->
<style>
#internal-mark-bar{{position:fixed;right:12px;bottom:10px;z-index:99999;
  font:12px/1.6 -apple-system,"Microsoft YaHei",sans-serif;color:#8a8a8a;
  background:rgba(255,255,255,.88);border:1px solid #e0e0e0;border-radius:4px;
  padding:2px 10px;pointer-events:none}}
@media print{{#internal-mark-bar{{position:fixed;bottom:6px}}}}
</style>
<div id="internal-mark-bar">{mark}</div>
"""


def stamp_html(path, mark, dry=False):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        s = f.read()
    if "internal-mark-bar" in s:
        return "html: 已有标注，跳过"
    block = HTML_BLOCK.format(mark=mark)
    if re.search(r"</body\s*>", s, re.I):
        s = re.sub(r"</body\s*>", block + "\n</body>", s, count=1, flags=re.I)
    else:
        s += block
    if not dry:
        with open(path, "w", encoding="utf-8") as f:
            f.write(s)
    return "html: 已加右下角标注条（含打印样式）"


# ---------------------------------------------------------------- pdf
def stamp_pdf(path, mark, dry=False):
    try:
        from pypdf import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas as rl_canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    except ImportError:
        return "pdf: 跳过（缺 pypdf/reportlab，建议在源 docx/pptx 上打标注后再导出）"

    import io
    fontname = "Helvetica"
    # 优先嵌入文泉驿（任何阅读器都能显示）；没有再用 STSong-Light（不嵌入，部分阅读器显示不出）
    for ttf in ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
                "/usr/local/share/fonts/company/wqy-zenhei.ttc"):
        if os.path.isfile(ttf):
            try:
                from reportlab.pdfbase.ttfonts import TTFont
                pdfmetrics.registerFont(TTFont("WQYZenHei", ttf, subfontIndex=0))
                fontname = "WQYZenHei"
                break
            except Exception:
                pass
    if fontname == "Helvetica":
        try:
            pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
            fontname = "STSong-Light"
        except Exception:
            pass

    reader = PdfReader(path)
    # 幂等：最后一页已有标注就跳过（重复运行不叠加）
    try:
        if mark in (reader.pages[-1].extract_text() or ""):
            return "pdf: 已有标注，跳过"
    except Exception:
        pass
    writer = PdfWriter()
    for page in reader.pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(w, h))
        c.setFont(fontname, 8)
        c.setFillColorRGB(0.5, 0.5, 0.5)
        c.drawRightString(w - 24, 16, mark)
        c.save()
        buf.seek(0)
        page.merge_page(PdfReader(buf).pages[0])
        writer.add_page(page)
    if not dry:
        with open(path, "wb") as f:
            writer.write(f)
    return f"pdf: {len(reader.pages)} 页加页脚标注"


# ---------------------------------------------------------------- driver
HANDLERS = {
    ".xlsx": stamp_xlsx, ".xlsm": stamp_xlsx,
    ".docx": stamp_docx,
    ".pptx": stamp_pptx,
    ".html": stamp_html, ".htm": stamp_html,
    ".pdf": stamp_pdf,
}
SKIP_PREFIX = ("~$", ".")


def iter_targets(paths, recursive):
    for p in paths:
        if os.path.isdir(p):
            walker = os.walk(p) if recursive else [(p, [], os.listdir(p))]
            for root, _dirs, files in walker:
                for fn in sorted(files):
                    if fn.startswith(SKIP_PREFIX):
                        continue
                    if os.path.splitext(fn)[1].lower() in HANDLERS:
                        yield os.path.join(root, fn)
        elif os.path.isfile(p):
            yield p


def main():
    ap = argparse.ArgumentParser(description="给交付文件加「内部文件，禁止外传」标注")
    ap.add_argument("paths", nargs="+", help="文件或目录")
    ap.add_argument("-r", "--recursive", action="store_true", help="目录递归")
    ap.add_argument("--text", default=MARK, help=f"标注文字（默认：{MARK}）")
    ap.add_argument("--dry-run", action="store_true", help="只报告不写回")
    a = ap.parse_args()

    ok = fail = 0
    for fp in iter_targets(a.paths, a.recursive):
        ext = os.path.splitext(fp)[1].lower()
        fn = HANDLERS.get(ext)
        if fn is None:
            print(f"[skip] {fp}（不支持 {ext}）")
            continue
        try:
            msg = fn(fp, a.text, a.dry_run)
            print(f"[ok]   {os.path.basename(fp)} — {msg}")
            ok += 1
        except Exception as e:
            print(f"[FAIL] {os.path.basename(fp)} — {type(e).__name__}: {e}")
            fail += 1
    print(f"\n完成：成功 {ok} 个，失败 {fail} 个"
          + ("（dry-run，未写回）" if a.dry_run else ""))
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
