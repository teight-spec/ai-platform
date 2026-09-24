#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
peek.py —— 数据文件快速摘要

处理任何 Excel/CSV 前先跑它，不要把整张表读进对话：
    python3 /opt/company/kit/peek.py 文件.xlsx                 # 所有 sheet 的概况
    python3 /opt/company/kit/peek.py 文件.xlsx --sheet 明细     # 只看一个 sheet
    python3 /opt/company/kit/peek.py 文件.xlsx --header 3       # 表头在第 3 行
    python3 /opt/company/kit/peek.py 文件.csv --rows 10         # 多看几行样例

输出（有上限，不会刷屏）：行数、列数、每列类型/空值/不同值个数/样例/数值范围/日期范围，前 N 行样例，
以及可能的问题（表头不在第 1 行、合并单元格导致的空列、文本型数字、重复行）。
"""
import argparse
import os
import sys
import warnings

warnings.filterwarnings("ignore")

MAX_COLS = 60
MAX_CELL = 18


def cut(v, n=MAX_CELL):
    if hasattr(v, "strftime"):
        v = v.strftime("%Y-%m-%d") if not getattr(v, "hour", 0) and not getattr(v, "minute", 0) else v.strftime("%Y-%m-%d %H:%M")
    s = str(v).replace("\n", " ")
    return s if len(s) <= n else s[: n - 1] + "…"


def nf(v):
    v = float(v)
    return f"{v:,.0f}" if abs(v) >= 1000 else f"{v:,.4g}"


def read_any(path, sheet, header):
    import pandas as pd
    ext = os.path.splitext(path)[1].lower()
    hdr = None if header == 0 else header - 1
    if ext in (".csv", ".txt"):
        for enc in ("utf-8-sig", "gbk", "gb18030"):
            try:
                return {"(CSV)": pd.read_csv(path, header=hdr, encoding=enc, dtype=object)}, enc
            except UnicodeDecodeError:
                continue
        raise SystemExit("CSV 编码识别失败（试过 utf-8 / gbk / gb18030）")
    engine = "xlrd" if ext == ".xls" else None
    data = pd.read_excel(path, sheet_name=sheet if sheet is not None else None, header=hdr, dtype=object, engine=engine)
    if not isinstance(data, dict):
        data = {sheet: data}
    return data, None


def guess_header(df_raw):
    """在前 10 行里找「非空文本最多」的一行，猜它是表头（1 起算）"""
    best, best_n = 1, -1
    for i in range(min(10, len(df_raw))):
        row = df_raw.iloc[i]
        n = sum(1 for v in row if isinstance(v, str) and v.strip())
        if n > best_n:
            best, best_n = i + 1, n
    return best


def col_profile(s):
    import pandas as pd
    nonnull = s.dropna()
    nonnull = nonnull[nonnull.astype(str).str.strip() != ""]
    info = {"空值": int(len(s) - len(nonnull)), "不同值": int(nonnull.astype(str).nunique())}
    if len(nonnull) == 0:
        info["类型"] = "全空"
        return info
    num = pd.to_numeric(nonnull.astype(str).str.replace(",", "").str.replace("，", "").str.strip(), errors="coerce")
    n_num = int(num.notna().sum())
    is_text_num = sum(1 for v in nonnull if isinstance(v, str)) > 0
    if n_num >= 0.9 * len(nonnull):
        info["类型"] = "数字" + ("（文本存储）" if is_text_num else "")
        info["范围"] = f"{nf(num.min())} ~ {nf(num.max())}，合计 {nf(num.sum())}"
        if n_num < len(nonnull):
            bad = nonnull[num.isna()].astype(str).unique()[:3]
            info["非数字"] = f"{len(nonnull) - n_num} 个，如 {', '.join(cut(b, 10) for b in bad)}"
        return info
    dt = pd.to_datetime(nonnull, errors="coerce")
    if dt.notna().sum() >= 0.9 * len(nonnull):
        info["类型"] = "日期"
        info["范围"] = f"{dt.min():%Y-%m-%d} ~ {dt.max():%Y-%m-%d}"
        if dt.isna().sum():
            info["无法识别"] = int(dt.isna().sum())
        return info
    info["类型"] = "文本"
    top = nonnull.astype(str).value_counts().head(3)
    info["常见值"] = "、".join(f"{cut(k, 12)}({v})" for k, v in top.items())
    return info


def main():
    ap = argparse.ArgumentParser(description="数据文件快速摘要")
    ap.add_argument("file")
    ap.add_argument("--sheet", default=None, help="只看这个 sheet（名称或从 0 开始的序号）")
    ap.add_argument("--header", type=int, default=1, help="表头所在行（1 起算；0 = 没有表头）")
    ap.add_argument("--rows", type=int, default=5, help="样例行数（最多 20）")
    a = ap.parse_args()
    import pandas as pd

    if not os.path.isfile(a.file):
        raise SystemExit(f"文件不存在：{a.file}")
    sheet = a.sheet
    if sheet is not None and sheet.isdigit():
        sheet = int(sheet)
    data, enc = read_any(a.file, sheet, a.header)
    size = os.path.getsize(a.file)
    print(f"文件：{a.file}（{size / 1024:,.0f} KB{'，编码 ' + enc if enc else ''}）")
    if len(data) > 1:
        print("Sheet 一览：" + "；".join(f"{k}（{len(v)} 行×{v.shape[1]} 列）" for k, v in data.items()))
    for name, df in data.items():
        print("\n" + "=" * 70)
        print(f"Sheet「{name}」：{len(df):,} 行 × {df.shape[1]} 列（表头取第 {a.header} 行）")
        df = df.dropna(how="all")
        issues = []
        unnamed = [c for c in df.columns if str(c).startswith("Unnamed")]
        if a.header and len(unnamed) >= max(2, df.shape[1] // 2):
            raw, _ = read_any(a.file, name if name != "(CSV)" else None, 0)
            g = guess_header(list(raw.values())[0])
            issues.append(f"表头可能不在第 {a.header} 行（{len(unnamed)} 列没有列名），疑似在第 {g} 行 → 加 --header {g} 再看")
        elif unnamed:
            issues.append(f"{len(unnamed)} 列没有列名（可能是合并单元格或备注列）：{', '.join(map(str, unnamed[:5]))}")
        dup = int(df.duplicated().sum())
        if dup:
            issues.append(f"完全重复的行 {dup} 行（是否去重要问用户）")
        cols = list(df.columns)[:MAX_COLS]
        print(f"{'列名':<16}{'类型':<12}{'空值':>6}{'不同值':>8}  说明")
        for c in cols:
            p = col_profile(df[c])
            extra = p.get("范围") or p.get("常见值") or ""
            for k in ("非数字", "无法识别"):
                if k in p:
                    extra += f"；{k}：{p[k]}"
            if "文本存储" in p.get("类型", ""):
                issues.append(f"「{cut(c)}」是文本存储的数字，计算前要转成数字")
            print(f"{cut(c, 15):<16}{p['类型']:<12}{p['空值']:>6}{p['不同值']:>8}  {extra}")
        if df.shape[1] > MAX_COLS:
            print(f"……还有 {df.shape[1] - MAX_COLS} 列未显示")
        n = max(1, min(a.rows, 20))
        print(f"\n前 {n} 行样例：")
        sample = df.head(n)[cols[:12]].copy()
        sample = sample.apply(lambda s: s.map(lambda v: "" if pd.isna(v) else cut(v, 14)))
        print(sample.to_string(index=False))
        if len(cols) > 12:
            print(f"（样例只显示前 12 列）")
        if issues:
            print("\n可能的问题：")
            for i in issues:
                print("  - " + i)
    print("\n提示：字段含义有疑问先问用户；正式计算写成 script.py，不要在对话里手算。")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:  # 给模型看得懂的报错
        print(f"读取失败：{type(e).__name__}: {e}")
        sys.exit(1)
