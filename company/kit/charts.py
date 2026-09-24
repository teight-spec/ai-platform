# -*- coding: utf-8 -*-
"""
charts.py —— 公司统一出图

用法（在终端的 Python 里）：
    import sys; sys.path.insert(0, "/opt/company/kit")
    from charts import line, bar, bar_compare, combo

    line(["W35","W36","W37"], {"入库金额": [12.1, 13.4, 11.8]}, "入库金额 · 近3周", "万元", "趋势.png")
    bar(["A仓","B仓"], [30, 12], "各仓库入库金额 · 2026-W38", "万元", "仓库.png")
    bar_compare(["A仓","B仓"], {"上周": [28, 15], "本周": [30, 12]}, "各仓库入库金额 · 本周 vs 上周", "万元", "对比.png")
    combo(["1月","2月"], [100, 120], [0.95, 1.02], "销售额与达成率", "万元", "达成率", "combo.png", rate=True)

规则（都已内置，不用自己设）：
  * 公司配色：蓝 2F6FB0、绿 3A9A5B、橙 D9822B、灰 8C8C8C
  * 类别超过 8 个自动改横向条形图并排序；数值标签自动加千分位
  * 标题写「什么数据 · 时间范围」，纵轴写单位
  * 输出 PNG 200dpi；返回保存路径
"""
import os
import warnings

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402

BLUE, GREEN, ORANGE, GRAY, DARK = "#2F6FB0", "#3A9A5B", "#D9822B", "#8C8C8C", "#262626"
PALETTE = [BLUE, GREEN, ORANGE, GRAY, "#7B5EA7", "#C0392B"]
_CJK = ["WenQuanYi Zen Hei", "Noto Sans CJK SC", "Microsoft YaHei", "SimHei", "DejaVu Sans"]


def _setup():
    warnings.filterwarnings("ignore", message=".*Glyph.*missing from font")  # 缺字警告会刷屏、浪费 token
    have = {f.name for f in font_manager.fontManager.ttflist}
    if not any(f in have for f in _CJK[:-1]):  # 系统里没有中文字体时，直接加载公司字体文件
        ttc = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fonts", "wqy-zenhei.ttc")
        if os.path.isfile(ttc):
            try:
                font_manager.fontManager.addfont(ttc)
                have = {f.name for f in font_manager.fontManager.ttflist}
            except Exception:
                pass
    fams = [f for f in _CJK if f in have] or ["DejaVu Sans"]
    plt.rcParams.update({
        "font.family": "sans-serif", "font.sans-serif": fams, "axes.unicode_minus": False,
        "axes.spines.top": False, "axes.spines.right": False, "axes.grid": True,
        "grid.color": "#EEEEEE", "axes.axisbelow": True, "axes.edgecolor": "#BFBFBF",
        "axes.titlesize": 13, "axes.titleweight": "bold", "axes.titlecolor": DARK,
        "axes.labelcolor": "#595959", "xtick.color": "#595959", "ytick.color": "#595959",
        "savefig.dpi": 200, "savefig.bbox": "tight", "figure.dpi": 110,
    })


_setup()


def fmt(v, decimals=None):
    """数值标签：千分位；大于 100 不带小数，否则 1 位（可指定）"""
    if v is None:
        return "—"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    if v != v:  # NaN
        return "—"
    if decimals is None:
        decimals = 0 if abs(v) >= 100 else (1 if abs(v) >= 1 else 2)
    return f"{v:,.{decimals}f}"


def pct(v, decimals=1):
    if v is None or v != v:
        return "—"
    return f"{v * 100:+.{decimals}f}%"


def _save(fig, out):
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    fig.savefig(out, facecolor="white")
    plt.close(fig)
    return out


def _nums(vals):
    return [None if v is None or (isinstance(v, float) and v != v) else float(v) for v in vals]


def line(x, series, title, unit, out, decimals=None, size=(8, 4.2), target=None, label_last_only=False):
    """折线图（看变化/趋势）。series = {"名称": [数值...]}，缺数据用 None。target=目标线数值。"""
    fig, ax = plt.subplots(figsize=size)
    x = [str(v) for v in x]
    for i, (name, vals) in enumerate(series.items()):
        vals = _nums(vals)
        c = PALETTE[i % len(PALETTE)]
        ax.plot(x, [float("nan") if v is None else v for v in vals], marker="o", lw=2.2, ms=5, color=c, label=name)
        for j, v in enumerate(vals):
            if v is None or (label_last_only and j != len(vals) - 1):
                continue
            last = j == len(vals) - 1
            ax.annotate(fmt(v, decimals), (x[j], v), textcoords="offset points", xytext=(0, 7), ha="center",
                        fontsize=9 if last else 8, color=c if last else "#595959", fontweight="bold" if last else None)
    if target is not None:
        ax.axhline(target, color=ORANGE, ls="--", lw=1.2)
        ax.annotate(f"目标 {fmt(target, decimals)}", (x[0], target), textcoords="offset points", xytext=(0, 4),
                    fontsize=8, color=ORANGE)
    ax.set_title(title, loc="left")
    ax.set_ylabel(unit or "")
    ax.margins(y=0.18)
    if len(series) > 1:
        ax.legend(frameon=False, ncol=min(len(series), 4), loc="upper left", bbox_to_anchor=(0, -0.1))
    return _save(fig, out)


def bar(categories, values, title, unit, out, decimals=None, sort=True, color=BLUE, size=None, top=None):
    """单系列柱状图（比较类别大小）。超过 8 类自动横向并排序；top=N 只画前 N 名。"""
    pairs = [(str(c), v) for c, v in zip(categories, _nums(values))]
    if sort:
        pairs.sort(key=lambda p: (p[1] is None, -(p[1] or 0)))
    if top:
        pairs = pairs[:top]
    cats = [p[0] for p in pairs]
    vals = [p[1] or 0 for p in pairs]
    horizontal = len(cats) > 8
    if horizontal:
        fig, ax = plt.subplots(figsize=size or (8, max(3.5, 0.38 * len(cats) + 1)))
        bars = ax.barh(cats[::-1], vals[::-1], color=color, height=0.6)
        for b, v in zip(bars, vals[::-1]):
            ax.annotate(fmt(v, decimals), (b.get_width(), b.get_y() + b.get_height() / 2), xytext=(4, 0),
                        textcoords="offset points", va="center", fontsize=8.5)
        ax.set_xlabel(unit or "")
        ax.grid(axis="y", visible=False)
        ax.margins(x=0.15)
    else:
        fig, ax = plt.subplots(figsize=size or (8, 4.2))
        bars = ax.bar(cats, vals, color=color, width=0.6)
        for b, v in zip(bars, vals):
            ax.annotate(fmt(v, decimals), (b.get_x() + b.get_width() / 2, b.get_height()), xytext=(0, 4),
                        textcoords="offset points", ha="center", fontsize=9)
        ax.set_ylabel(unit or "")
        ax.grid(axis="x", visible=False)
        ax.margins(y=0.15)
        if max((len(c) for c in cats), default=0) > 5:
            plt.setp(ax.get_xticklabels(), rotation=20, ha="right")
    ax.set_title(title, loc="left")
    return _save(fig, out)


def bar_compare(categories, series, title, unit, out, decimals=None, size=None):
    """多系列柱状图（如 上周 vs 本周）。series = {"上周": [...], "本周": [...]}，最后一个系列用蓝色突出。"""
    cats = [str(c) for c in categories]
    names = list(series.keys())
    n = len(names)
    colors = ([GRAY] * (n - 1) + [BLUE]) if n <= 2 else PALETTE[:n]
    horizontal = len(cats) > 8
    import numpy as np
    idx = np.arange(len(cats))
    w = 0.8 / n
    if horizontal:
        fig, ax = plt.subplots(figsize=size or (8, max(3.8, 0.5 * len(cats) + 1)))
        for i, name in enumerate(names):
            vals = [v or 0 for v in _nums(series[name])]
            pos = idx[::-1] + (n - 1 - i) * w - 0.4 + w / 2
            bs = ax.barh(pos, vals, height=w, color=colors[i], label=name)
            for b, v in zip(bs, vals):
                ax.annotate(fmt(v, decimals), (b.get_width(), b.get_y() + b.get_height() / 2), xytext=(3, 0),
                            textcoords="offset points", va="center", fontsize=7.5)
        ax.set_yticks(idx[::-1])
        ax.set_yticklabels(cats)
        ax.set_xlabel(unit or "")
        ax.grid(axis="y", visible=False)
        ax.margins(x=0.15)
    else:
        fig, ax = plt.subplots(figsize=size or (8, 4.2))
        for i, name in enumerate(names):
            vals = [v or 0 for v in _nums(series[name])]
            bs = ax.bar(idx + i * w - 0.4 + w / 2, vals, width=w, color=colors[i], label=name)
            for b, v in zip(bs, vals):
                ax.annotate(fmt(v, decimals), (b.get_x() + b.get_width() / 2, b.get_height()), xytext=(0, 3),
                            textcoords="offset points", ha="center", fontsize=7.5)
        ax.set_xticks(idx)
        ax.set_xticklabels(cats)
        if max((len(c) for c in cats), default=0) > 5:
            plt.setp(ax.get_xticklabels(), rotation=20, ha="right")
        ax.set_ylabel(unit or "")
        ax.grid(axis="x", visible=False)
        ax.margins(y=0.15)
    ax.set_title(title, loc="left", pad=14)
    ax.legend(frameon=False, ncol=n, loc="lower right", bbox_to_anchor=(1, 1.0), fontsize=9)
    return _save(fig, out)


def combo(x, bars_vals, line_vals, title, bar_unit, line_label, out, rate=False, decimals=None, size=(8, 4.2)):
    """柱+折线双轴（如 金额 + 达成率）。rate=True 时折线按百分比显示。"""
    x = [str(v) for v in x]
    fig, ax = plt.subplots(figsize=size)
    bvals = [v or 0 for v in _nums(bars_vals)]
    bs = ax.bar(x, bvals, color=BLUE, width=0.6, label=bar_unit)
    for b, v in zip(bs, bvals):
        ax.annotate(fmt(v, decimals), (b.get_x() + b.get_width() / 2, b.get_height()), xytext=(0, 3),
                    textcoords="offset points", ha="center", fontsize=8.5)
    ax.set_ylabel(bar_unit or "")
    ax.grid(axis="x", visible=False)
    ax.margins(y=0.2)
    ax2 = ax.twinx()
    lv = _nums(line_vals)
    ax2.plot(x, [float("nan") if v is None else v for v in lv], color=ORANGE, marker="o", lw=2, label=line_label)
    for xi, v in zip(x, lv):
        if v is not None:
            ax2.annotate(f"{v * 100:.1f}%" if rate else fmt(v, decimals), (xi, v), xytext=(0, 7),
                         textcoords="offset points", ha="center", fontsize=8.5, color=ORANGE)
    ax2.set_ylabel(line_label)
    ax2.spines["right"].set_visible(True)
    ax2.grid(False)
    ax2.margins(y=0.25)
    if rate:
        ax2.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v * 100:.0f}%"))
    ax.set_title(title, loc="left")
    return _save(fig, out)
