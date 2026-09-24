#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dept.py —— 部门说明 + 部门配方

让助手「越用越准、越用越省」：确认过的口径记进部门说明，做对过的任务存成配方，下次直接重跑。

    python3 /opt/company/kit/dept.py start                         任务开始先跑：部门说明 + 配方清单（一次看全）
    python3 /opt/company/kit/dept.py note --section 字段口径 --by 张三 "入库金额 = 数量 × 含税单价，不含退货"
    python3 /opt/company/kit/dept.py recipe-save --name 月度入库汇总 --from <任务输出目录> --desc "一句话说明" --by 张三
    python3 /opt/company/kit/dept.py recipe-list
    python3 /opt/company/kit/dept.py recipe-run 月度入库汇总 --input 新文件.xlsx [--user 张三] [--out 目录]
    python3 /opt/company/kit/dept.py recipe-check [配方名 …]        用样例重跑、和验收值比对（升级/改工具包后跑）

所有子命令都可加 --dept-dir <部门文件夹>；不加时自动找（当前目录往上找含 04_技能 的文件夹，或 /home/user 下唯一的部门文件夹）。
"""
import argparse
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runio  # noqa: E402

MARK = "内部文件，禁止外传"
NOTES = os.path.join("05_模板", "部门说明.md")
WEEKLY = os.path.join("05_模板", "周报说明.md")
SKILLS = "04_技能"
SECTIONS = ["术语", "字段口径", "常用文件与来源", "汇报习惯", "其他"]
MAX_SAMPLE_MB = 30
MAX_PRINT_LINES = 120


# ----------------------------------------------------------------- 部门文件夹
def find_dept_dir(given=None):
    if given:
        d = os.path.abspath(os.path.expanduser(given))
        if not os.path.isdir(d):
            raise SystemExit(f"部门文件夹不存在：{d}")
        return d
    env = os.getenv("AI_DEPT_DIR")
    if env and os.path.isdir(env):
        return os.path.abspath(env)
    d = os.getcwd()
    while d and d != os.path.dirname(d):
        if os.path.isdir(os.path.join(d, SKILLS)) and os.path.isdir(os.path.join(d, "05_模板")):
            return d
        d = os.path.dirname(d)
    home = os.path.expanduser("~")
    cands = [os.path.join(home, n) for n in sorted(os.listdir(home)) if os.path.isdir(os.path.join(home, n, SKILLS))] \
        if os.path.isdir(home) else []
    if len(cands) == 1:
        return cands[0]
    raise SystemExit("找不到部门文件夹，请加 --dept-dir ~/<部门>")


def today():
    return dt.date.today().strftime("%Y-%m-%d")


# ----------------------------------------------------------------- 部门说明
def notes_template(dept_name):
    body = [f"# {dept_name}部门说明", "", f"<!-- {MARK} -->",
            "助手每次任务前会读本文件。只写和用户确认过的内容，每条注明日期和确认人；保持简短（建议 100 行以内）。",
            "要修改或删除某条，直接编辑本文件即可。", ""]
    for s in SECTIONS:
        body += [f"## {s}", ""]
    return "\n".join(body)


def count_items(text):
    return sum(1 for ln in text.splitlines() if ln.lstrip().startswith("- "))


def cmd_note(a):
    root = find_dept_dir(a.dept_dir)
    if a.section not in SECTIONS:
        raise SystemExit(f"--section 只能是：{'、'.join(SECTIONS)}")
    text = " ".join(a.text).strip().replace("\n", " ")
    if not text:
        raise SystemExit("内容为空")
    p = os.path.join(root, NOTES)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    doc = open(p, encoding="utf-8").read() if os.path.isfile(p) else notes_template(os.path.basename(root))
    if re.search(r"^- " + re.escape(text) + r"（", doc, re.M):
        print(f"已有相同条目，未重复写入：{text}")
        return
    line = f"- {text}（{today()} {a.by or '未注明'}确认）"
    lines = doc.splitlines()
    head = f"## {a.section}"
    if head not in lines:
        lines += ["", head, ""]
    i = lines.index(head) + 1
    j = i
    while j < len(lines) and not lines[j].startswith("## "):
        j += 1
    k = j  # 插在本节最后一条非空行之后
    while k > i and not lines[k - 1].strip():
        k -= 1
    lines.insert(k, line)
    out = "\n".join(lines).rstrip() + "\n"
    with open(p, "w", encoding="utf-8") as f:
        f.write(out)
    print(f"已写入 {NOTES}「{a.section}」：{line}")
    n = count_items(out)
    if len(out.splitlines()) > 100:
        print(f"提示：部门说明已有 {len(out.splitlines())} 行（{n} 条），建议请用户合并精简，太长会增加每次对话的费用。")


# ----------------------------------------------------------------- 配方
def recipes(root):
    base = os.path.join(root, SKILLS)
    out = []
    if not os.path.isdir(base):
        return out
    for n in sorted(os.listdir(base)):
        d = os.path.join(base, n)
        if n.startswith(".") or not os.path.isfile(os.path.join(d, "script.py")) or not os.path.isfile(os.path.join(d, "SKILL.md")):
            continue
        meta = {"name": n, "dir": d, "desc": "", "checked": ""}
        m = re.search(r"^description:\s*(.+)$", open(os.path.join(d, "SKILL.md"), encoding="utf-8").read(), re.M)
        if m:
            meta["desc"] = m.group(1).strip()
        acc = os.path.join(d, "验收.json")
        if os.path.isfile(acc):
            try:
                meta["checked"] = json.load(open(acc, encoding="utf-8")).get("_验收时间", "")
            except Exception:
                pass
        out.append(meta)
    return out


def other_skills(root):
    base = os.path.join(root, SKILLS)
    if not os.path.isdir(base):
        return []
    rs = {r["name"] for r in recipes(root)}
    return [n for n in sorted(os.listdir(base))
            if n not in rs and not n.startswith(".") and os.path.isfile(os.path.join(base, n, "SKILL.md"))]


def cmd_start(a):
    root = find_dept_dir(a.dept_dir)
    print(f"部门文件夹：{root}")
    p = os.path.join(root, NOTES)
    if os.path.isfile(p):
        text = open(p, encoding="utf-8").read()
        lines, head = [], None
        for ln in text.splitlines():  # 只打印有内容的小节（标题下的说明文字不打印），省 token
            if ln.startswith("## "):
                head = ln
            elif head is not None and ln.strip() and not ln.strip().startswith("<!--"):
                if head:
                    lines.append(head)
                    head = ""
                lines.append(ln)
        print(f"\n【部门说明】{NOTES}（{count_items(text)} 条，以下口径已和用户确认过，照此执行）")
        for ln in lines[:MAX_PRINT_LINES]:
            print(ln)
        if len(lines) > MAX_PRINT_LINES:
            print(f"……（还有 {len(lines) - MAX_PRINT_LINES} 行，需要时用 cat 看全文；建议请用户精简）")
    else:
        print(f"\n【部门说明】还没有（{NOTES}）。任务中和用户确认的口径，结束时用 note 记下来。")
    rs = recipes(root)
    print(f"\n【部门配方】{len(rs)} 个" + ("（同类任务先用配方：recipe-run）" if rs else "：暂无。任务做对后可以存成配方（recipe-save）。"))
    for r in rs:
        print(f"  - {r['name']}：{r['desc'] or '（无说明）'}" + (f"　[验收 {r['checked'][:10]}]" if r["checked"] else ""))
    others = other_skills(root)
    if others:
        print("【部门技能】" + "、".join(others) + f"（在 {SKILLS}/<名称>/SKILL.md，和公司技能冲突时以部门为准）")
    print("\n【周报说明】" + ("已设置：" + WEEKLY if os.path.isfile(os.path.join(root, WEEKLY)) else "未设置（第一次做周报用 /周报设置）"))


def _read_summary(d):
    p = os.path.join(d, runio.SUMMARY)
    if not os.path.isfile(p):
        return None
    return json.load(open(p, encoding="utf-8"))


def _values(doc):
    return {k: v for k, v in (doc or {}).items() if not k.startswith("_")}


def _diff(expect, got, path=""):
    """比较两份结果摘要的数字（相对误差 1e-6）。返回不一致的描述列表"""
    out = []
    if isinstance(expect, dict) and isinstance(got, dict):
        for k in expect:
            if k not in got:
                out.append(f"{path}{k}：缺少（应为 {expect[k]}）")
            else:
                out += _diff(expect[k], got[k], f"{path}{k}.")
        for k in got:
            if k not in expect:
                out.append(f"{path}{k}：多出（{got[k]}）")
        return out
    if isinstance(expect, (int, float)) and isinstance(got, (int, float)) and not isinstance(expect, bool):
        if abs(expect - got) > 1e-6 * max(1.0, abs(expect)):
            out.append(f"{path.rstrip('.')}：应为 {expect}，实际 {got}")
        return out
    if expect != got:
        out.append(f"{path.rstrip('.')}：应为 {expect!r}，实际 {got!r}")
    return out


def _run_script(script, inputs, out, cwd):
    os.makedirs(out, exist_ok=True)
    r = subprocess.run([sys.executable, script, "--input", *inputs, "--out", out], cwd=cwd,
                       capture_output=True, text=True, timeout=900)
    return r


def _tail(text, n=30):
    lines = (text or "").strip().splitlines()
    return "\n".join(lines[-n:])


def _structure_diff(sample_struct, new_file):
    """新文件和样例的 sheet / 表头差异。返回问题列表（空 = 一致）"""
    new = runio.structure(new_file)
    if not sample_struct or not new:
        return []
    probs = []
    for sheet, cols in sample_struct.items():
        if sheet not in new:
            if len(sample_struct) == 1 and len(new) == 1:  # 只有一个 sheet 时允许改名
                new_cols = next(iter(new.values()))
            else:
                probs.append(f"缺少 sheet「{sheet}」")
                continue
        else:
            new_cols = new[sheet]
        miss = [c for c in cols if c not in new_cols]
        if miss:
            probs.append(f"sheet「{sheet}」缺少列：{'、'.join(miss[:10])}")
        extra = [c for c in new_cols if c not in cols]
        if extra:
            probs.append(f"sheet「{sheet}」多了列：{'、'.join(extra[:10])}（多列一般不影响，但请确认口径没变）")
    return probs


def cmd_recipe_save(a):
    root = find_dept_dir(a.dept_dir)
    name = a.name.strip()
    if not name or re.search(r"[\\/:*?\"<>|]", name):
        raise SystemExit("配方名不能为空，也不能含 \\ / : * ? \" < > |")
    src = os.path.abspath(os.path.expanduser(a.from_dir))
    script = os.path.join(src, "script.py")
    if not os.path.isfile(script):
        raise SystemExit(f"{src} 里没有 script.py")
    code = open(script, encoding="utf-8").read()
    if "io_args" not in code or "save_summary" not in code:
        raise SystemExit("script.py 要先改成可重跑的写法：用 runio.io_args 接收 --input/--out，用 save_summary 写结果摘要"
                         "（见 /opt/company/skills/公司-部门说明与配方/SKILL.md），改好重跑一次再存。")
    summary = _read_summary(src)
    if summary is None:
        raise SystemExit(f"{src} 里没有 {runio.SUMMARY}：请先运行一次 script.py")
    inputs = [os.path.abspath(os.path.expanduser(p)) for p in (a.input or summary.get("_输入路径") or [])]
    missing = [p for p in inputs if not os.path.isfile(p)]
    if not inputs or missing:
        raise SystemExit("找不到样例输入文件：" + ("；".join(missing) if missing else "请用 --input 指定"))
    dst = os.path.join(root, SKILLS, name)
    if os.path.exists(dst) and not a.overwrite:
        raise SystemExit(f"配方「{name}」已存在。确认要更新就加 --overwrite（旧版会移到 {name}/历史/）")
    big = [p for p in inputs if os.path.getsize(p) > MAX_SAMPLE_MB * 1024 * 1024]
    if big:
        raise SystemExit(f"样例文件超过 {MAX_SAMPLE_MB}MB，不适合存进配方：{'；'.join(big)}。请用一份小一点的样例重跑后再存。")
    # 先在临时文件夹里建好并验收，通过了再换上去；失败时原配方不受影响
    stage = os.path.join(root, SKILLS, f".{name}.新建中")
    shutil.rmtree(stage, ignore_errors=True)
    os.makedirs(os.path.join(stage, "样例"))
    shutil.copy2(script, os.path.join(stage, "script.py"))
    samples = []
    for p in inputs:
        q = os.path.join(stage, "样例", os.path.basename(p))
        shutil.copy2(p, q)
        samples.append(q)
    tmp = tempfile.mkdtemp(prefix="配方验收_")
    r = _run_script(os.path.join(stage, "script.py"), samples, tmp, root)
    got = _read_summary(tmp)
    shutil.rmtree(tmp, ignore_errors=True)
    if r.returncode != 0 or got is None:
        shutil.rmtree(stage, ignore_errors=True)
        raise SystemExit("用样例重跑失败，配方没有保存（脚本可能依赖了别的文件）：\n" + _tail(r.stderr or r.stdout))
    diff = _diff(_values(summary), _values(got))
    if diff:
        print("注意：用样例重跑的结果和原任务不一致（脚本可能依赖了别的文件或手工步骤）：")
        for d in diff[:10]:
            print("  - " + d)
        print("已按重跑结果作为验收值保存；请和用户确认哪个是对的。")
    acc = _values(got)
    acc["_验收时间"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    acc["_样例"] = [os.path.basename(p) for p in samples]
    acc["_样例结构"] = got.get("_输入结构", {})
    acc["_说明"] = MARK
    with open(os.path.join(stage, "验收.json"), "w", encoding="utf-8") as f:
        json.dump(acc, f, ensure_ascii=False, indent=2)
    notes = [n.strip() for n in (a.note or []) if n.strip()]
    skill = [
        "---", f"name: {name}", f"description: {a.desc.strip()}", "---", f"<!-- {MARK} -->", "",
        f"# {name}（部门配方）", "",
        f"- 用途：{a.desc.strip()}",
        f"- 建立：{today()}，{a.by or '未注明'}确认；来源：{os.path.relpath(src, root)}",
        f"- 输入：{'、'.join(os.path.basename(p) for p in samples)}（新文件的 sheet 和列名要和样例一致，样例在 样例/ 里）",
        "", "## 怎么用",
        f"1. `python3 /opt/company/kit/dept.py recipe-run \"{name}\" --input <新文件> --user <用户名>`",
        "   新文件和样例格式不一致时会提示，先停下来问用户，确认后再改脚本（改完用 recipe-save --overwrite 重新存）。",
        "2. 结论只引用输出目录里 结果摘要.json 的数字；需要出图、PPT 时照公司技能做。",
        "3. 交付前照常运行 stamp_internal.py 加标注。",
    ]
    if notes:
        skill += ["", "## 注意事项"] + [f"- {n}" for n in notes]
    with open(os.path.join(stage, "SKILL.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(skill) + "\n")
    if os.path.exists(dst):  # 更新：旧版整体移到 历史/<时间>/
        hist = os.path.join(stage, "历史")
        if os.path.isdir(os.path.join(dst, "历史")):
            shutil.move(os.path.join(dst, "历史"), hist)
        os.makedirs(hist, exist_ok=True)
        shutil.move(dst, os.path.join(hist, dt.datetime.now().strftime("%Y%m%d_%H%M%S")))
    os.rename(stage, dst)
    print(f"已保存部门配方「{name}」：{os.path.relpath(dst, root)}/（script.py、SKILL.md、验收.json、样例/）")
    print("验收值：" + json.dumps({k: v for k, v in acc.items() if not k.startswith("_")}, ensure_ascii=False)[:300])


def _find_recipe(root, name):
    for r in recipes(root):
        if r["name"] == name:
            return r
    names = "、".join(r["name"] for r in recipes(root)) or "（暂无）"
    raise SystemExit(f"没有配方「{name}」。现有配方：{names}")


def cmd_recipe_list(a):
    root = find_dept_dir(a.dept_dir)
    rs = recipes(root)
    if not rs:
        print("暂无部门配方。")
        return
    for r in rs:
        print(f"- {r['name']}：{r['desc'] or '（无说明）'}" + (f"　[验收 {r['checked'][:10]}]" if r["checked"] else ""))


def cmd_recipe_run(a):
    root = find_dept_dir(a.dept_dir)
    r = _find_recipe(root, a.name)
    inputs = [os.path.abspath(os.path.expanduser(p)) for p in a.input]
    missing = [p for p in inputs if not os.path.exists(p)]
    if missing:
        raise SystemExit("找不到输入文件：" + "；".join(missing))
    acc = {}
    try:
        acc = json.load(open(os.path.join(r["dir"], "验收.json"), encoding="utf-8"))
    except Exception:
        pass
    sample_struct = acc.get("_样例结构") or {}
    if sample_struct and not a.force:
        probs = []
        samples = list(sample_struct.values())
        for i, p in enumerate(inputs):
            if i < len(samples):
                probs += [f"{os.path.basename(p)}：{x}" for x in _structure_diff(samples[i], p)]
        hard = [x for x in probs if "缺少" in x]
        for x in probs:
            print("格式检查：" + x)
        if hard:
            raise SystemExit("新文件和样例格式不一致，没有运行。请先和用户确认：是文件换了格式，还是放错了文件？"
                             "确认口径没变可以加 --force 强制运行。")
    if a.out:
        out = os.path.abspath(os.path.expanduser(a.out))
    else:
        stamp = dt.date.today().strftime("%Y%m%d")
        base = os.path.join(root, "02_输出", f"{stamp}_{a.user + '_' if a.user else ''}{a.name}")
        out, n = base, 2
        while os.path.exists(out):
            out, n = f"{base}_{n}", n + 1
    res = _run_script(os.path.join(r["dir"], "script.py"), inputs, out, root)
    if res.returncode != 0:
        raise SystemExit("配方运行失败：\n" + _tail(res.stderr or res.stdout))
    shutil.copy2(os.path.join(r["dir"], "script.py"), os.path.join(out, "script.py"))
    with open(os.path.join(out, "配方来源.txt"), "w", encoding="utf-8") as f:
        f.write(f"{MARK}\n配方：{r['name']}（{os.path.relpath(r['dir'], root)}）\n运行时间：{dt.datetime.now():%Y-%m-%d %H:%M:%S}\n"
                f"输入：{'；'.join(inputs)}\n")
    print(_tail(res.stdout, 20))
    got = _read_summary(out)
    print(f"\n输出目录：{out}")
    if got:
        print("结果摘要：" + json.dumps(_values(got), ensure_ascii=False)[:600])
    else:
        print(f"注意：脚本没有写 {runio.SUMMARY}，结论里的数字请从输出文件核对。")


def cmd_recipe_check(a):
    root = find_dept_dir(a.dept_dir)
    rs = [_find_recipe(root, n) for n in a.names] if a.names else recipes(root)
    if not rs:
        print("暂无部门配方，无需检查。")
        return 0
    bad = 0
    for r in rs:
        try:
            acc = json.load(open(os.path.join(r["dir"], "验收.json"), encoding="utf-8"))
        except Exception:
            print(f"[跳过] {r['name']}：没有 验收.json")
            continue
        samples = [os.path.join(r["dir"], "样例", n) for n in acc.get("_样例", [])]
        if not samples or not all(os.path.isfile(p) for p in samples):
            print(f"[跳过] {r['name']}：样例文件不全")
            continue
        tmp = tempfile.mkdtemp(prefix="配方回归_")
        res = _run_script(os.path.join(r["dir"], "script.py"), samples, tmp, root)
        got = _read_summary(tmp)
        if res.returncode != 0 or got is None:
            bad += 1
            print(f"[失败] {r['name']}：运行出错\n" + _tail(res.stderr or res.stdout, 8))
        else:
            diff = _diff(_values(acc), _values(got))
            if diff:
                bad += 1
                print(f"[不一致] {r['name']}：" + "；".join(diff[:5]))
            else:
                print(f"[OK] {r['name']}")
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"共 {len(rs)} 个配方，{'全部通过' if not bad else f'{bad} 个有问题'}")
    return 1 if bad else 0


def main():
    ap = argparse.ArgumentParser(description="部门说明 + 部门配方")
    ap.add_argument("--dept-dir", default=None, help="部门文件夹（默认自动找）")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("start", help="任务开始：显示部门说明 + 配方清单")
    s = sub.add_parser("note", help="往部门说明里记一条确认过的口径")
    s.add_argument("--section", required=True, help="、".join(SECTIONS))
    s.add_argument("--by", default="", help="确认人")
    s.add_argument("text", nargs="+")
    s = sub.add_parser("recipe-save", help="把做对的任务存成部门配方")
    s.add_argument("--name", required=True)
    s.add_argument("--from", dest="from_dir", required=True, help="任务输出目录（里面有 script.py 和 结果摘要.json）")
    s.add_argument("--desc", required=True, help="一句话说明：做什么、用什么文件")
    s.add_argument("--by", default="", help="确认人")
    s.add_argument("--input", nargs="+", help="样例输入文件（默认用结果摘要里记录的输入）")
    s.add_argument("--note", nargs="*", help="注意事项，可多条")
    s.add_argument("--overwrite", action="store_true")
    s = sub.add_parser("recipe-list", help="列出部门配方")
    s = sub.add_parser("recipe-run", help="用配方处理新文件")
    s.add_argument("name")
    s.add_argument("--input", nargs="+", required=True)
    s.add_argument("--out", default=None)
    s.add_argument("--user", default="", help="用户名（用于输出目录命名）")
    s.add_argument("--force", action="store_true", help="格式和样例不一致也运行（先和用户确认）")
    s = sub.add_parser("recipe-check", help="用样例重跑所有配方并比对验收值")
    s.add_argument("names", nargs="*")
    # 子命令后面也接受 --dept-dir
    for p in sub.choices.values():
        p.add_argument("--dept-dir", default=argparse.SUPPRESS, help=argparse.SUPPRESS)
    a = ap.parse_args()
    fn = {"start": cmd_start, "note": cmd_note, "recipe-save": cmd_recipe_save, "recipe-list": cmd_recipe_list,
          "recipe-run": cmd_recipe_run, "recipe-check": cmd_recipe_check}[a.cmd]
    sys.exit(fn(a) or 0)


if __name__ == "__main__":
    main()
