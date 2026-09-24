#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
term_guard.py —— 部门终端守护

由 term_entry.sh 在部门终端启动时放到后台运行，只用 Python 标准库，做两件事：
  1. 超时命令终止：助手通过终端执行的命令（Open Terminal 的 /execute，进程形如 /bin/sh -c ...）
     运行超过 TERM_CMD_MAX_MIN 分钟（默认 20）就连同它的子进程一起结束（先 SIGTERM，10 秒后 SIGKILL）。
     Open Terminal 自己不会结束卡住的命令，一个死循环会一直占着部门终端的 CPU 和内存。
     员工在网页「终端」面板里手动打开的交互式 shell 不受影响。
  2. 临时文件清理：每小时清一次 /tmp 下超过 TERM_TMP_KEEP_HOURS 小时（默认 24）没动过的临时文件和目录
     （助手按规定用 mktemp -d 建临时目录；部门终端长期不重建，残留会越积越多）。

日志直接打到容器日志（Container Manager → 容器 → 日志），每次终止或清理都有一行记录。
手工检查：python3 /opt/company/term_guard.py --once   （只扫描一次并打印，不常驻）
"""
import os
import shutil
import signal
import sys
import time
from datetime import datetime

MAX_MIN = float(os.getenv("TERM_CMD_MAX_MIN", "20"))
TMP_KEEP_H = float(os.getenv("TERM_TMP_KEEP_HOURS", "24"))
SCAN_SEC = float(os.getenv("TERM_GUARD_SCAN_SEC", "30"))
TMP_EVERY_SEC = 3600
TMP_DIR = os.getenv("TERM_GUARD_TMP", "/tmp")
KEEP_IN_TMP = {".X11-unix", ".ICE-unix", ".font-unix", ".XIM-unix", ".Test-unix"}
HZ = os.sysconf("SC_CLK_TCK")


def log(msg):
    print(f"[term-guard {datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def proc_table():
    """{pid: (ppid, pgid, 启动秒数(开机后), argv)}"""
    out = {}
    for d in os.listdir("/proc"):
        if not d.isdigit():
            continue
        try:
            with open(f"/proc/{d}/stat", "rb") as f:
                st = f.read().decode("utf-8", "replace")
            rest = st[st.rindex(")") + 2:].split()
            ppid, pgid, start = int(rest[1]), int(rest[2]), int(rest[19]) / HZ
            with open(f"/proc/{d}/cmdline", "rb") as f:
                argv = [a.decode("utf-8", "replace") for a in f.read().split(b"\0") if a]
            out[int(d)] = (ppid, pgid, start, argv)
        except (OSError, ValueError, IndexError):
            continue
    return out


def uptime():
    with open("/proc/uptime") as f:
        return float(f.read().split()[0])


def descendants(table, root):
    kids = {}
    for pid, (ppid, *_rest) in table.items():
        kids.setdefault(ppid, []).append(pid)
    out, stack = [], [root]
    while stack:
        for c in kids.get(stack.pop(), []):
            out.append(c)
            stack.append(c)
    return out


def long_commands(table, server_pid, now):
    """Open Terminal 直接起的「/bin/sh -c 命令」且运行超时的：[(pid, 分钟, 命令)]"""
    res = []
    for pid, (ppid, _pgid, start, argv) in table.items():
        if ppid != server_pid or pid == os.getpid():
            continue
        if len(argv) < 3 or argv[1] != "-c" or not argv[0].endswith("sh"):
            continue  # 交互式 shell（终端面板）、notebook 内核等不管
        minutes = (now - start) / 60
        if minutes > MAX_MIN:
            res.append((pid, minutes, argv[2]))
    return res


def stop_tree(table, pid):
    victims = [pid] + descendants(table, pid)
    pgids = {table[p][1] for p in victims if p in table}
    for sig in (signal.SIGTERM, signal.SIGKILL):
        for g in pgids:
            try:
                os.killpg(g, sig)
            except (ProcessLookupError, PermissionError):
                pass
        for p in victims:  # 子进程自己 setsid 另起进程组的（比如 soffice），逐个补刀
            try:
                os.kill(p, sig)
            except (ProcessLookupError, PermissionError):
                pass
        if sig == signal.SIGTERM:
            for _ in range(20):
                time.sleep(0.5)
                if not any(os.path.exists(f"/proc/{p}") and _alive(p) for p in victims):
                    return len(victims)
    return len(victims)


def _alive(pid):
    try:
        with open(f"/proc/{pid}/stat", "rb") as f:
            st = f.read().decode("utf-8", "replace")
        return st[st.rindex(")") + 2:].split()[0] != "Z"  # 僵尸进程算已结束
    except OSError:
        return False


def clean_tmp(now_ts):
    uid = os.getuid()
    n, freed = 0, 0
    for name in os.listdir(TMP_DIR):
        if name in KEEP_IN_TMP:
            continue
        p = os.path.join(TMP_DIR, name)
        try:
            st = os.lstat(p)
        except OSError:
            continue
        if st.st_uid != uid or now_ts - st.st_mtime < TMP_KEEP_H * 3600:
            continue
        try:
            if os.path.isdir(p) and not os.path.islink(p):
                size = sum(os.path.getsize(os.path.join(r, f)) for r, _d, fs in os.walk(p) for f in fs
                           if os.path.isfile(os.path.join(r, f)))
                shutil.rmtree(p, ignore_errors=True)
            else:
                size = st.st_size
                os.remove(p)
            n += 1
            freed += size
        except OSError:
            continue
    if n:
        log(f"清理 /tmp：删除 {n} 个超过 {TMP_KEEP_H:g} 小时的临时文件/目录，释放 {freed / 1024 / 1024:.1f} MB")
    return n


def scan(server_pid):
    table = proc_table()
    now = uptime()
    killed = 0
    for pid, minutes, cmd in long_commands(table, server_pid, now):
        n = stop_tree(table, pid)
        killed += 1
        log(f"终止超时命令（已运行 {minutes:.1f} 分钟，上限 {MAX_MIN:g} 分钟，连同子进程共 {n} 个）：{cmd[:200]}")
    return killed


def main():
    if "--once" in sys.argv:
        server = int(sys.argv[sys.argv.index("--server") + 1]) if "--server" in sys.argv else None
        table = proc_table()
        if server is None:  # 找 open-terminal 主进程
            server = next((p for p, v in table.items() if any("open-terminal" in a for a in v[3][:2])), None)
        print(f"Open Terminal 主进程：{server}；命令上限 {MAX_MIN:g} 分钟；/tmp 保留 {TMP_KEEP_H:g} 小时")
        now = uptime()
        for pid, (ppid, _g, start, argv) in sorted(table.items()):
            if ppid == server and len(argv) >= 3 and argv[1] == "-c":
                print(f"  运行中 pid={pid} 已 {(now - start) / 60:.1f} 分钟：{argv[2][:120]}")
        return
    server_pid = os.getppid()  # term_entry.sh 用 exec 交给原版入口，父进程就是 Open Terminal 主进程
    log(f"启动：守护 Open Terminal（pid {server_pid}），命令超过 {MAX_MIN:g} 分钟自动终止，/tmp 保留 {TMP_KEEP_H:g} 小时")
    last_tmp = 0.0
    while True:
        time.sleep(SCAN_SEC)
        if not _alive(server_pid):
            log("Open Terminal 主进程已退出，守护结束")
            return
        try:
            scan(server_pid)
            if time.time() - last_tmp >= TMP_EVERY_SEC:
                last_tmp = time.time()
                clean_tmp(last_tmp)
        except Exception as e:  # 守护自己出错不能影响终端
            log(f"出错（已忽略）：{type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
