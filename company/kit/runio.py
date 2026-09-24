# -*- coding: utf-8 -*-
"""
runio.py —— 让 script.py 能「换个文件直接重跑」

每个任务的 script.py 开头都这样写（默认值 = 本次用的文件和输出目录）：
    import sys; sys.path.insert(0, "/opt/company/kit")
    from runio import io_args, save_summary
    inputs, out = io_args(["~/资材/01_原始数据/入库明细_0923.xlsx"], "~/资材/02_输出/20260924_张三_入库汇总")
    ...计算...
    save_summary({"入库总金额(元)": 1234567.8, "有效行数": 5210, "按仓库": {"A仓": 1.2, "B仓": 3.4}})

    重跑：python3 script.py --input 新文件.xlsx --out 新输出目录

save_summary 在输出目录写「结果摘要.json」：结论里引用的关键数字 + 输入文件 + 输入结构（sheet 和表头），
用来：① 结论里的数字可追溯；② 存成部门配方后，dept.py 用它做回归比对、检查新文件格式是否和样例一致。
"""
import argparse
import datetime as _dt
import json
import os
import sys

_STATE = {"inputs": [], "out": None}
SUMMARY = "结果摘要.json"
MARK = "内部文件，禁止外传"


def _abs(p):
    return os.path.abspath(os.path.expanduser(p))


def io_args(default_inputs=None, default_out=None):
    """解析 --input（可多个）和 --out；没给就用默认值。返回 (输入文件列表, 输出目录)，输出目录会自动建好。"""
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--input", nargs="+")
    ap.add_argument("--out")
    a, _ = ap.parse_known_args()
    if isinstance(default_inputs, str):
        default_inputs = [default_inputs]
    inputs = [_abs(p) for p in (a.input or default_inputs or [])]
    out = _abs(a.out or default_out or os.getcwd())
    if not inputs:
        raise SystemExit("缺少输入文件：用 --input 指定")
    missing = [p for p in inputs if not os.path.exists(p)]
    if missing:
        raise SystemExit("找不到输入文件：" + "；".join(missing))
    os.makedirs(out, exist_ok=True)
    _STATE.update(inputs=inputs, out=out)
    return inputs, out


def _header_guess(df_raw):
    """前 10 行里非空文本最多的一行当表头（和 peek.py 同一规则）"""
    best, best_n = 0, -1
    for i in range(min(10, len(df_raw))):
        n = sum(1 for v in df_raw.iloc[i] if isinstance(v, str) and v.strip())
        if n > best_n:
            best, best_n = i, n
    return best


def structure(path):
    """文件结构指纹：{sheet 名: [表头列名…]}。只读前 10 行，大文件也很快；不是表格文件返回 {}"""
    import warnings
    warnings.filterwarnings("ignore")
    ext = os.path.splitext(path)[1].lower()
    try:
        import pandas as pd
        if ext in (".csv", ".txt"):
            data = None
            for enc in ("utf-8-sig", "gbk", "gb18030"):
                try:
                    data = {"(CSV)": pd.read_csv(path, header=None, nrows=10, dtype=object, encoding=enc)}
                    break
                except UnicodeDecodeError:
                    continue
            if data is None:
                return {}
        elif ext in (".xlsx", ".xlsm", ".xls"):
            data = pd.read_excel(path, sheet_name=None, header=None, nrows=10, dtype=object,
                                 engine="xlrd" if ext == ".xls" else None)
        else:
            return {}
    except Exception:
        return {}
    out = {}
    for name, df in data.items():
        if df.empty:
            out[str(name)] = []
            continue
        row = df.iloc[_header_guess(df)]
        out[str(name)] = [str(v).strip() for v in row if isinstance(v, str) and v.strip()]
    return out


def _plain(v):
    """numpy / pandas 的数字、日期转成 JSON 能存的普通值"""
    if isinstance(v, dict):
        return {str(k): _plain(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [_plain(x) for x in v]
    if hasattr(v, "item") and not isinstance(v, (str, bytes)):
        try:
            return v.item()
        except Exception:
            pass
    if hasattr(v, "isoformat"):
        return v.isoformat()
    if isinstance(v, float) and v != v:
        return None
    return v


def save_summary(values, out=None):
    """写 结果摘要.json。values = 结论里要引用的关键数字（可以嵌套一层字典）。返回文件路径。"""
    out = _abs(out) if out else _STATE["out"]
    if not out:
        raise SystemExit("save_summary：先调用 io_args()，或传 out=输出目录")
    os.makedirs(out, exist_ok=True)
    doc = {k: _plain(v) for k, v in values.items()}
    doc["_输入"] = [os.path.basename(p) for p in _STATE["inputs"]]
    doc["_输入路径"] = _STATE["inputs"]
    doc["_输入结构"] = {os.path.basename(p): structure(p) for p in _STATE["inputs"]}
    doc["_脚本"] = _abs(sys.argv[0]) if sys.argv and sys.argv[0] else ""
    doc["_生成时间"] = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    doc["_说明"] = MARK
    p = os.path.join(out, SUMMARY)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    return p
