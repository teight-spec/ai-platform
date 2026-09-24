# -*- coding: utf-8 -*-
"""
费用闸门 cost-gate —— 公司内部 AI 平台的「唯一出网点」
============================================================
作用：
  1. 对 Open WebUI 提供 OpenAI 兼容接口，原样转发到智谱开放平台
  2. 按「部门密钥」识别部门，调用前检查本月额度，超限拦截
  3. 按实际 usage 计价（人民币），逐条记录调用日志（可追溯）
  4. 自检页  http://NAS:8031/        连通性、密钥、额度概览
  5. 管理页  http://NAS:8031/admin   费用明细、改上限、导出 CSV（需管理员密码）
     月度价值  http://NAS:8031/admin/value   活跃人数、成果数、每个成果的费用、部门说明与配方的积累
     配置检查  管理页按钮：和一键初始化时的基线比较，发现有人在 Open WebUI 界面上手工改过配置
  6. 模拟模式 GATE_MOCK=1（只在本机测试用）：不连智谱，返回假回复 + 按字数估算的费用，
     整条链路（登录、部门权限、终端工具调用、计费、限额拦截）都能在没有 API key 时测试

运行环境：直接复用 Open WebUI 镜像里的 Python（fastapi / uvicorn / httpx 已内置），
          不需要额外构建镜像、不需要联网装包。
"""
import asyncio
import base64
import csv
import hmac
import html
import io
import json
import math
import os
import re
import sqlite3
import threading
import time
import urllib.parse
from datetime import datetime, timedelta, timezone

import httpx
import uvicorn
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response, StreamingResponse, RedirectResponse

VERSION = "1.4.0"
CN_TZ = timezone(timedelta(hours=8))  # 中国不用夏令时，固定 UTC+8


# ----------------------------------------------------------------- 配置
def _env(name, default=""):
    v = os.getenv(name, default).strip()
    return "" if v.startswith("请填写") else v  # .env 里未改的占位符视为未配置


UPSTREAM_BASE = _env("UPSTREAM_BASE", "https://open.bigmodel.cn/api/paas/v4").rstrip("/")
UPSTREAM_KEY = _env("ZHIPU_API_KEY")
ADMIN_PASSWORD = _env("GATE_ADMIN_PASSWORD")
DB_PATH = _env("GATE_DB_PATH", "/data/gate.db")
MODELS = [m.strip() for m in _env("GATE_MODELS", "glm-5.3-flash").split(",") if m.strip()]
PRICE_IN = float(_env("PRICE_INPUT", "0.8"))       # 元 / 百万 tokens（未命中缓存的输入）
PRICE_OUT = float(_env("PRICE_OUTPUT", "2.8"))     # 元 / 百万 tokens（输出，含思考）
PRICE_CACHE = float(_env("PRICE_CACHED", "0.23"))  # 元 / 百万 tokens（命中缓存的输入）
PRICE_EMBED = float(_env("PRICE_EMBEDDING", "0.5"))
WARN_PCT = float(_env("WARN_PCT", "80"))
CHARS_PER_TOKEN = float(_env("EST_CHARS_PER_TOKEN", "1.5"))  # 上游没回 usage 时的估算口径
STRIP_STREAM_OPTIONS = _env("STRIP_STREAM_OPTIONS", "true").lower() == "true"
OWUI_URL = _env("OWUI_URL", "http://open-webui:8080")
GATE_INTERNAL_URL = _env("GATE_INTERNAL_URL", "http://cost-gate:8000")
EXEC_DEPT = _env("GATE_EXEC_DEPT", "zb")  # 总裁办：助手只读各部门文件夹，账号为管理员
VIEWER_PASSWORD = _env("GATE_VIEWER_PASSWORD")  # 只读密码：只能看费用和记录，不能改
MOCK = _env("GATE_MOCK", "0").lower() in ("1", "true", "yes", "on")  # 模拟模式：只在本机测试时打开
MOCK_DELAY = float(_env("GATE_MOCK_DELAY_MS", "25")) / 1000          # 模拟流式输出每段间隔
MOCK_NOTE = "模拟模式（未连接智谱，费用为估算）"
# 单个对话的累计费用上限（元）：防止一个对话来回调用终端把部门额度吃光；0 = 不限
CHAT_MAX_YUAN = float(_env("CHAT_MAX_YUAN", "5"))
# 单次请求的输入上限（估算 tokens）：对话太长或把大表格读进对话时拦下来；0 = 不限
MAX_INPUT_TOKENS = int(float(_env("MAX_INPUT_TOKENS", "100000")))
# 工具输出截断：单条工具结果（终端命令输出、读文件）超过这么多字时，只把开头和结尾发给模型（截掉的部分员工在界面上仍能看到）；0 = 不截断
TOOL_OUTPUT_MAX_CHARS = int(float(_env("TOOL_OUTPUT_MAX_CHARS", "12000")))
# 长对话提示：本对话上一次输入超过这么多 tokens，或本对话费用超过单对话上限的这个百分比时，提示员工新开对话；0 = 不提示
CHAT_HINT_TOKENS = int(float(_env("CHAT_HINT_TOKENS", "60000")))
CHAT_HINT_PCT = float(_env("CHAT_HINT_PCT", "60"))
# 部门说明篇幅上限（字）与复核周期（天）：和终端里 dept.py 一致，月度价值页据此标黄
NOTES_MAX_CHARS = int(float(_env("AI_NOTES_MAX_CHARS", "3000")))
NOTES_REVIEW_DAYS = int(float(_env("NOTES_REVIEW_DAYS", "31")))
# 月度价值页：部门共享文件夹只读挂载在这里（只统计 02_输出、03_周报、04_技能、05_模板 的目录名和条目数，不读文件内容）
DEPTS_ROOT = _env("DEPTS_ROOT", "/depts")
# 一键初始化后的 Open WebUI 配置基线（配置检查用）
BASELINE_DIR = os.path.join(os.path.dirname(DB_PATH) or "/data", "owui_config")
try:
    EXTRA_BODY = json.loads(_env("GATE_EXTRA_BODY", "{}") or "{}")
except Exception:
    EXTRA_BODY = {}


def parse_depts(raw):
    """GATE_DEPTS 格式：代码|名称|月上限元|密钥;代码|名称|月上限元|密钥"""
    out = {}
    for part in raw.split(";"):
        part = part.strip()
        if not part:
            continue
        f = [x.strip() for x in part.split("|")]
        if len(f) != 4 or not f[3]:
            raise SystemExit(f"GATE_DEPTS 配置格式错误：{part!r}（应为 代码|名称|月上限|密钥）")
        out[f[0]] = {"code": f[0], "name": f[1], "limit": float(f[2]), "key": f[3]}
    return out


DEPTS = parse_depts(_env("GATE_DEPTS"))
KEY_TO_DEPT = {d["key"]: code for code, d in DEPTS.items()}
DEPT_COLORS = ["#2F6FB0", "#3A9A5B", "#D9822B", "#8C8C8C", "#7B5EA7", "#C0392B"]


# ----------------------------------------------------------------- 数据库
_db_lock = threading.Lock()


def db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    with _db_lock, db() as c:
        c.execute("PRAGMA journal_mode=WAL")
        c.execute("""CREATE TABLE IF NOT EXISTS calls(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT, month TEXT, day TEXT,
            dept TEXT, user_id TEXT, user_name TEXT, user_email TEXT, user_role TEXT,
            chat_id TEXT, task TEXT, endpoint TEXT, model TEXT, stream INTEGER,
            prompt_tokens INTEGER, completion_tokens INTEGER, cached_tokens INTEGER,
            cost REAL, estimated INTEGER, status INTEGER, latency_ms INTEGER, note TEXT)""")
        c.execute("CREATE INDEX IF NOT EXISTS ix_calls_dm ON calls(dept, month)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_calls_ts ON calls(ts)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_calls_chat ON calls(chat_id)")
        c.execute("""CREATE TABLE IF NOT EXISTS budgets(
            dept TEXT PRIMARY KEY, limit_yuan REAL, updated_at TEXT)""")
        for code, d in DEPTS.items():  # 首次启动写入初始上限；之后以管理页修改为准
            c.execute("INSERT OR IGNORE INTO budgets(dept, limit_yuan, updated_at) VALUES(?,?,?)",
                      (code, d["limit"], now_cn().isoformat(timespec="seconds")))


def now_cn():
    return datetime.now(CN_TZ)


def month_of(dt=None):
    return (dt or now_cn()).strftime("%Y-%m")


def get_limit(dept):
    with db() as c:
        r = c.execute("SELECT limit_yuan FROM budgets WHERE dept=?", (dept,)).fetchone()
    return float(r["limit_yuan"]) if r else DEPTS.get(dept, {}).get("limit", 0.0)


def get_spent(dept, month=None):
    with db() as c:
        r = c.execute("SELECT COALESCE(SUM(cost),0) s FROM calls WHERE dept=? AND month=?",
                      (dept, month or month_of())).fetchone()
    return float(r["s"])


def get_chat_spent(chat_id):
    if not chat_id:
        return 0.0
    with db() as c:
        r = c.execute("SELECT COALESCE(SUM(cost),0) s FROM calls WHERE chat_id=?", (chat_id,)).fetchone()
    return float(r["s"])


def get_chat_last_input(chat_id):
    """本对话最近一次成功的主对话调用的输入 tokens（标题生成等后台调用不算）"""
    if not chat_id:
        return 0
    with db() as c:
        r = c.execute("SELECT prompt_tokens FROM calls WHERE chat_id=? AND task='chat' AND status=200 "
                      "ORDER BY id DESC LIMIT 1", (chat_id,)).fetchone()
    return int(r["prompt_tokens"] or 0) if r else 0


def chat_hint(chat_spent, last_input):
    """长对话提示：返回提示文字，或空字符串"""
    why = []
    if CHAT_HINT_TOKENS and last_input >= CHAT_HINT_TOKENS:
        why.append(f"上一轮输入约 {last_input / 10000:.1f} 万 tokens")
    if CHAT_MAX_YUAN > 0 and CHAT_HINT_PCT and chat_spent >= CHAT_MAX_YUAN * CHAT_HINT_PCT / 100:
        why.append(f"已用 ¥{chat_spent:.2f}（上限 ¥{CHAT_MAX_YUAN:g}）")
    if not why:
        return ""
    return "这个对话已经比较长（" + "，".join(why) + "），每轮都更费钱。建议做完这一步就点「新对话」：已确认的口径在部门说明里，文件都在部门文件夹，不会丢。"


def spend_info(dept, month=None):
    month = month or month_of()
    spent, limit = get_spent(dept, month), get_limit(dept)
    pct = (spent / limit * 100) if limit > 0 else 100.0  # 上限填 0 = 暂停该部门使用
    return {"dept": dept, "name": DEPTS[dept]["name"], "month": month,
            "spent": round(spent, 4), "limit": round(limit, 2), "pct": round(pct, 1),
            "warn": pct >= WARN_PCT, "blocked": spent >= limit}


def log_call(**kw):
    ts = now_cn()
    row = dict(ts=ts.isoformat(timespec="seconds"), month=month_of(ts), day=ts.strftime("%Y-%m-%d"))
    row.update(kw)
    cols = ",".join(row.keys())
    qs = ",".join("?" for _ in row)
    with _db_lock, db() as c:
        c.execute(f"INSERT INTO calls({cols}) VALUES({qs})", list(row.values()))


def calc_cost(prompt, completion, cached):
    uncached = max(prompt - cached, 0)
    return (uncached * PRICE_IN + cached * PRICE_CACHE + completion * PRICE_OUT) / 1_000_000


# ----------------------------------------------------------------- 工具函数
from contextlib import asynccontextmanager

_client: httpx.AsyncClient = None
STARTED = now_cn()


@asynccontextmanager
async def lifespan(_app):
    global _client
    init_db()
    _client = httpx.AsyncClient(timeout=httpx.Timeout(connect=10, read=900, write=60, pool=30))
    yield
    await _client.aclose()


app = FastAPI(title="cost-gate", docs_url=None, redoc_url=None, openapi_url=None, lifespan=lifespan)


def err(status, message, typ="gate_error"):
    return JSONResponse({"error": {"message": message, "type": typ}}, status_code=status)


def who(request: Request):
    h = request.headers
    return dict(
        user_id=h.get("x-openwebui-user-id", ""),
        user_name=urllib.parse.unquote(h.get("x-openwebui-user-name", "")),
        user_email=h.get("x-openwebui-user-email", ""),
        user_role=h.get("x-openwebui-user-role", ""),
        chat_id=h.get("x-openwebui-chat-id", ""),
        task=h.get("x-owui-task", "") or "chat",
    )


def auth_dept(request: Request):
    auth = request.headers.get("authorization", "")
    key = auth[7:].strip() if auth.lower().startswith("bearer ") else ""
    for k, code in KEY_TO_DEPT.items():
        if key and hmac.compare_digest(k, key):
            return code
    return None


def usage_tuple(u):
    if not isinstance(u, dict):
        return None
    p = int(u.get("prompt_tokens") or 0)
    c = int(u.get("completion_tokens") or 0)
    details = u.get("prompt_tokens_details") or {}
    cached = int((details.get("cached_tokens") if isinstance(details, dict) else 0) or 0)
    return p, c, cached


def est_tokens(text_len):
    return int(text_len / CHARS_PER_TOKEN) + 1


def upstream_headers():
    return {"Authorization": f"Bearer {UPSTREAM_KEY}", "Content-Type": "application/json"}


TRUNC_NOTE = ("……【费用闸门：这段工具输出共 {n} 字，太长，中间 {cut} 字没有发给模型（界面上仍能看到全文）。"
              "需要细节时：把结果写进文件，再用 head / grep / peek.py 只看需要的部分】……")


def _cut(text):
    if len(text) <= TOOL_OUTPUT_MAX_CHARS or "【费用闸门：这段工具输出共" in text:
        return text, False
    head = int(TOOL_OUTPUT_MAX_CHARS * 0.6)
    tail = TOOL_OUTPUT_MAX_CHARS - head
    return text[:head] + "\n" + TRUNC_NOTE.format(n=len(text), cut=len(text) - head - tail) + "\n" + text[-tail:], True


def truncate_tool_outputs(body):
    """把过长的工具结果（role=tool）截成「开头 60% + 结尾 40%」。规则固定，同一条历史每次截得一样，不影响缓存命中。
    返回截断的条数。"""
    if not TOOL_OUTPUT_MAX_CHARS:
        return 0
    n = 0
    for m in body.get("messages") or []:
        if not isinstance(m, dict) or m.get("role") != "tool":
            continue
        c = m.get("content")
        if isinstance(c, str):
            m["content"], hit = _cut(c)
            n += hit
        elif isinstance(c, list):
            for part in c:
                if isinstance(part, dict) and isinstance(part.get("text"), str):
                    part["text"], hit = _cut(part["text"])
                    n += hit
    return n


def check_limits(dept, meta, prompt_chars):
    """单对话费用上限 + 单次输入上限。返回 None=放行，或 (状态码, 类型, 提示)"""
    if MAX_INPUT_TOKENS and est_tokens(prompt_chars) > MAX_INPUT_TOKENS:
        return (413, "input_too_long",
                f"本次发送的内容约 {est_tokens(prompt_chars) / 10000:.1f} 万 tokens，超过单次上限 {MAX_INPUT_TOKENS / 10000:g} 万。"
                f"通常是对话太长，或把整张大表格读进了对话。请点「新对话」重新开始；大文件请让助手先用 peek.py 出摘要。")
    chat_id = meta.get("chat_id")
    if CHAT_MAX_YUAN > 0 and chat_id and not chat_id.startswith("local:"):
        spent = get_chat_spent(chat_id)
        if spent >= CHAT_MAX_YUAN:
            return (402, "chat_budget_exceeded",
                    f"这个对话已用 ¥{spent:.2f}，达到单个对话上限 ¥{CHAT_MAX_YUAN:g}。为了节约部门额度，请点「新对话」重新开始"
                    f"（需要的结论可以复制过去，文件都在部门文件夹里）。")
    return None


def budget_block_message(info):
    return (f"【{info['name']}】本月 AI 额度已用完：已用 ¥{info['spent']:.2f} / 上限 ¥{info['limit']:.2f}。"
            f"请联系平台管理员调整额度，下月 1 日自动恢复。")


# ----------------------------------------------------------------- 模拟模式（本机测试用）
def _text_of(content):
    """OpenAI 消息的 content 可能是字符串，也可能是 [{type:text,text:..}, {type:image_url..}]"""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        return "".join(p.get("text", "") for p in content if isinstance(p, dict))
    return ""


def _tool_names(body):
    out = []
    for t in body.get("tools") or []:
        fn = (t or {}).get("function") or {}
        if fn.get("name"):
            out.append(fn)
    return out


def mock_reply(body, task):
    """返回 (content, reasoning, tool_calls)。规则固定、可预期，方便对照测试。"""
    fixed = {"title_generation": '{"title": "模拟对话"}', "tags_generation": '{"tags": ["模拟"]}',
             "follow_up_generation": '{"follow_ups": []}', "query_generation": '{"queries": []}',
             "autocomplete_generation": '{"text": ""}', "emoji_generation": "🙂"}
    if task in fixed:
        return fixed[task], "", None
    msgs = body.get("messages") or []
    last = msgs[-1] if msgs else {}
    tools = _tool_names(body)
    names = [f["name"] for f in tools]
    if last.get("role") == "tool":
        out = _text_of(last.get("content"))
        cut = "（发给模型前已被费用闸门截断）" if "【费用闸门：这段工具输出共" in out else ""
        return (f"【模拟回复】工具已执行，返回 {len(out)} 字{cut}，前 800 字如下：\n\n```\n{out[:800]}\n```\n\n"
                f"（链路正常：模型 → 工具 → 结果回传。）"), "模拟思考：读取工具结果。", None
    text = _text_of(last.get("content")).strip()
    nudged = "（平台提示，不用单独回复" in text
    text = text.split("\n\n（平台提示，不用单独回复")[0].strip()  # 过滤器加的长对话收尾提示，不影响模拟命令
    if text.startswith("#工具"):
        if not tools:
            return "【模拟回复】本次请求没有带任何工具。请确认对话里已选择部门终端（或部门助手已绑定终端）。", "", None
        lines = []
        for f in tools:
            props = ((f.get("parameters") or {}).get("properties") or {})
            lines.append(f"- `{f['name']}`（参数：{', '.join(props) or '无'}）：{(f.get('description') or '')[:60]}")
        return "【模拟回复】本次请求可用的工具：\n" + "\n".join(lines) + \
            '\n\n调用方法：发送 `#调用 工具名 {"参数名": "值"}`', "", None
    if text.startswith("#调用"):
        m = re.match(r"#调用\s+(\S+)\s*(.*)$", text, re.S)
        if not m:
            return '【模拟回复】格式：`#调用 工具名 {"参数名": "值"}`', "", None
        name, args = m.group(1), (m.group(2).strip() or "{}")
        if name not in names:
            return f"【模拟回复】找不到工具 `{name}`。本次可用：{', '.join(names) or '无'}", "", None
        try:
            json.loads(args)
        except Exception:
            return f"【模拟回复】参数不是合法 JSON：`{args[:200]}`", "", None
        call = {"id": "call_mock_" + str(int(time.time() * 1000)), "type": "function",
                "function": {"name": name, "arguments": args}}
        return "", f"模拟思考：按指令调用工具 {name}。", [call]
    nudge = "\n- 消息末尾带有长对话收尾提示（过滤器已生效）" if nudged else ""
    return (f"【模拟回复 · 未连接智谱】收到你的消息（{len(text)} 字）。{nudge}\n\n"
            f"- 模型：{body.get('model', '')}\n- 本次消息条数：{len(msgs)}\n"
            f"- 可用工具：{len(names)} 个{('（' + '、'.join(names[:8]) + '）') if names else ''}\n\n"
            "测试命令：发送 `#工具` 查看可用工具；发送 `#调用 工具名 {\"参数\": \"值\"}` 让模拟模型调用工具。"), \
        "模拟思考：这是固定的假回复，用于检查界面和计费。", None


def _chunks(text, size=8):
    return [text[i:i + size] for i in range(0, len(text), size)] if text else []


async def mock_chat(body, dept, model, stream, meta, t0, cut_note=""):
    content, reasoning, tool_calls = mock_reply(body, meta.get("task", "chat"))
    prompt_chars = len(json.dumps(body.get("messages", []), ensure_ascii=False)) + \
        len(json.dumps(body.get("tools", []), ensure_ascii=False))
    out_chars = len(content) + len(reasoning) + len(json.dumps(tool_calls or [], ensure_ascii=False))
    p, c = est_tokens(prompt_chars), est_tokens(out_chars)
    usage = {"prompt_tokens": p, "completion_tokens": c, "total_tokens": p + c,
             "prompt_tokens_details": {"cached_tokens": 0}}
    finish = "tool_calls" if tool_calls else "stop"
    cid, created = f"mock-{int(time.time() * 1000)}", int(time.time())

    def log(note=MOCK_NOTE):
        note = "；".join(x for x in (note, cut_note) if x)
        log_call(dept=dept, endpoint="chat", model=model, stream=int(stream), prompt_tokens=p, completion_tokens=c,
                 cached_tokens=0, cost=calc_cost(p, c, 0), estimated=1, status=200,
                 latency_ms=int((time.monotonic() - t0) * 1000), note=note, **meta)

    if not stream:
        msg = {"role": "assistant", "content": content or None}
        if reasoning:
            msg["reasoning_content"] = reasoning
        if tool_calls:
            msg["tool_calls"] = tool_calls
        log()
        return JSONResponse({"id": cid, "object": "chat.completion", "created": created, "model": model,
                             "choices": [{"index": 0, "message": msg, "finish_reason": finish}], "usage": usage})

    def sse(delta=None, finish_reason=None, usage_obj=None):
        obj = {"id": cid, "object": "chat.completion.chunk", "created": created, "model": model,
               "choices": [] if usage_obj else [{"index": 0, "delta": delta or {}, "finish_reason": finish_reason}]}
        if usage_obj:
            obj["usage"] = usage_obj
        return ("data: " + json.dumps(obj, ensure_ascii=False) + "\n\n").encode("utf-8")

    async def gen():
        note = MOCK_NOTE
        try:
            yield sse({"role": "assistant", "content": ""})
            for piece in _chunks(reasoning):
                yield sse({"reasoning_content": piece})
                await asyncio.sleep(MOCK_DELAY)
            for piece in _chunks(content):
                yield sse({"content": piece})
                await asyncio.sleep(MOCK_DELAY)
            for i, tc in enumerate(tool_calls or []):
                yield sse({"tool_calls": [dict(index=i, **tc)]})
            yield sse({}, finish)
            yield sse(usage_obj=usage)
            yield b"data: [DONE]\n\n"
        except asyncio.CancelledError:
            note = MOCK_NOTE + "；用户中途停止"
            raise
        finally:
            try:
                log(note)
            except Exception:
                pass

    return StreamingResponse(gen(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


def mock_embeddings(body):
    import hashlib
    import math
    inp = body.get("input", "")
    items = inp if isinstance(inp, list) else [inp]
    dim = int(body.get("dimensions") or 1024)
    data, chars = [], 0
    for i, it in enumerate(items):
        t = it if isinstance(it, str) else json.dumps(it)
        chars += len(t)
        seed = hashlib.sha256(t.encode("utf-8")).digest()
        vec, k = [], 0
        while len(vec) < dim:  # 同一段文字永远得到同一个向量，方便复测
            block = hashlib.sha256(seed + k.to_bytes(4, "little")).digest()
            vec.extend((b - 127.5) / 127.5 for b in block)
            k += 1
        vec = vec[:dim]
        n = math.sqrt(sum(v * v for v in vec)) or 1.0
        data.append({"object": "embedding", "index": i, "embedding": [round(v / n, 6) for v in vec]})
    p = est_tokens(chars)
    return {"object": "list", "data": data, "model": body.get("model", ""),
            "usage": {"prompt_tokens": p, "total_tokens": p}}


# ----------------------------------------------------------------- OpenAI 兼容接口
@app.get("/v1/models")
async def models():
    return {"object": "list", "data": [{"id": m, "object": "model", "owned_by": "zhipu"} for m in MODELS]}


@app.post("/v1/chat/completions")
async def chat_completions(request: Request):
    dept = auth_dept(request)
    if not dept:
        return err(401, "费用闸门：部门密钥无效，请检查 Open WebUI 连接配置。", "invalid_key")
    if not UPSTREAM_KEY and not MOCK:
        return err(503, "费用闸门：尚未配置智谱 API 密钥（.env 里的 ZHIPU_API_KEY）。", "no_upstream_key")
    info = spend_info(dept)
    if info["blocked"]:
        log_call(dept=dept, endpoint="chat", model="", stream=0, prompt_tokens=0, completion_tokens=0, cached_tokens=0,
                 cost=0.0, estimated=0, status=402, latency_ms=0, note="部门本月额度已用完", **who(request))
        return err(402, budget_block_message(info), "budget_exceeded")

    try:
        body = await request.json()
    except Exception:
        return err(400, "费用闸门：请求体不是合法 JSON。")
    for k, v in EXTRA_BODY.items():
        body.setdefault(k, v)
    if STRIP_STREAM_OPTIONS:
        body.pop("stream_options", None)
    stream = bool(body.get("stream"))
    model = str(body.get("model", ""))
    meta = who(request)
    n_cut = truncate_tool_outputs(body)
    cut_note = f"截断工具输出 {n_cut} 处" if n_cut else ""
    prompt_chars = len(json.dumps(body.get("messages", []), ensure_ascii=False)) + \
        len(json.dumps(body.get("tools", []), ensure_ascii=False))
    t0 = time.monotonic()
    blocked = check_limits(dept, meta, prompt_chars)
    if blocked:
        status, typ, message = blocked
        log_call(dept=dept, endpoint="chat", model=model, stream=int(stream), prompt_tokens=0,
                 completion_tokens=0, cached_tokens=0, cost=0.0, estimated=0, status=status,
                 latency_ms=0, note=message[:120], **meta)
        return err(status, message, typ)
    if MOCK:
        return await mock_chat(body, dept, model, stream, meta, t0, cut_note)
    url = f"{UPSTREAM_BASE}/chat/completions"

    try:
        req = _client.build_request("POST", url, headers=upstream_headers(), json=body)
        resp = await _client.send(req, stream=True)
    except Exception as e:
        log_call(dept=dept, endpoint="chat", model=model, stream=int(stream), prompt_tokens=0,
                 completion_tokens=0, cached_tokens=0, cost=0.0, estimated=0, status=599,
                 latency_ms=int((time.monotonic() - t0) * 1000), note=f"上游连接失败: {type(e).__name__}", **meta)
        return err(502, f"费用闸门：连接智谱失败（{type(e).__name__}）。请打开自检页查看网络状态。", "upstream_unreachable")

    if resp.status_code != 200 or not stream:
        raw = await resp.aread()
        await resp.aclose()
        latency = int((time.monotonic() - t0) * 1000)
        p = c = cached = 0
        estimated = 0
        note = ""
        try:
            data = json.loads(raw)
        except Exception:
            data = None
        if resp.status_code == 200 and isinstance(data, dict):
            ut = usage_tuple(data.get("usage"))
            if ut:
                p, c, cached = ut
            else:
                estimated = 1
                out_chars = len(json.dumps(data.get("choices", []), ensure_ascii=False))
                p, c = est_tokens(prompt_chars), est_tokens(out_chars)
        elif resp.status_code != 200:
            note = raw[:300].decode("utf-8", "replace")
        if resp.status_code == 200:
            note = "；".join(x for x in (note, cut_note) if x)
        cost = calc_cost(p, c, cached) if resp.status_code == 200 else 0.0
        log_call(dept=dept, endpoint="chat", model=model, stream=int(stream), prompt_tokens=p,
                 completion_tokens=c, cached_tokens=cached, cost=cost, estimated=estimated,
                 status=resp.status_code, latency_ms=latency, note=note, **meta)
        ctype = resp.headers.get("content-type", "application/json")
        return Response(content=raw, status_code=resp.status_code, media_type=ctype.split(";")[0])

    # ---- 流式：边转发边解析最后一个 chunk 里的 usage
    state = {"usage": None, "out_chars": 0, "buf": b"", "done": False}

    def scan(chunk: bytes):
        state["buf"] += chunk
        while b"\n" in state["buf"]:
            line, state["buf"] = state["buf"].split(b"\n", 1)
            line = line.strip()
            if not line.startswith(b"data:"):
                continue
            payload = line[5:].strip()
            if payload == b"[DONE]":
                state["done"] = True
                continue
            try:
                obj = json.loads(payload)
            except Exception:
                continue
            ut = usage_tuple(obj.get("usage"))
            if ut:
                state["usage"] = ut
            for ch in obj.get("choices") or []:
                d = ch.get("delta") or {}
                for key in ("content", "reasoning_content"):
                    if isinstance(d.get(key), str):
                        state["out_chars"] += len(d[key])
                for tc in d.get("tool_calls") or []:
                    fn = (tc or {}).get("function") or {}
                    state["out_chars"] += len(fn.get("arguments") or "") + len(fn.get("name") or "")

    async def relay():
        note = ""
        try:
            async for chunk in resp.aiter_raw():
                scan(chunk)
                yield chunk
        except asyncio.CancelledError:
            note = "用户中途停止"
            raise
        except Exception as e:
            note = f"流中断: {type(e).__name__}"
            raise
        finally:
            await resp.aclose()
            if state["buf"]:
                scan(b"\n")
            if state["usage"]:
                p, c, cached = state["usage"]
                estimated = 0
            else:
                p, c, cached = est_tokens(prompt_chars), est_tokens(state["out_chars"]), 0
                estimated = 1
                note = note or "上游未返回 usage，按字符估算"
            note = "；".join(x for x in (note, cut_note) if x)
            try:
                log_call(dept=dept, endpoint="chat", model=model, stream=1, prompt_tokens=p,
                         completion_tokens=c, cached_tokens=cached, cost=calc_cost(p, c, cached),
                         estimated=estimated, status=200, latency_ms=int((time.monotonic() - t0) * 1000),
                         note=note, **meta)
            except Exception:
                pass

    headers = {"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}
    return StreamingResponse(relay(), media_type="text/event-stream", headers=headers)


@app.post("/v1/embeddings")
async def embeddings(request: Request):
    dept = auth_dept(request)
    if not dept:
        return err(401, "费用闸门：部门密钥无效。", "invalid_key")
    info = spend_info(dept)
    if info["blocked"]:
        return err(402, budget_block_message(info), "budget_exceeded")
    body = await request.json()
    meta = who(request)
    t0 = time.monotonic()
    if MOCK:
        data = mock_embeddings(body)
        p = data["usage"]["prompt_tokens"]
        meta["task"] = "embedding"
        log_call(dept=dept, endpoint="embeddings", model=str(body.get("model", "")), stream=0, prompt_tokens=p,
                 completion_tokens=0, cached_tokens=0, cost=p * PRICE_EMBED / 1_000_000, estimated=1, status=200,
                 latency_ms=int((time.monotonic() - t0) * 1000), note=MOCK_NOTE, **meta)
        return JSONResponse(data)
    try:
        resp = await _client.post(f"{UPSTREAM_BASE}/embeddings", headers=upstream_headers(), json=body)
    except Exception as e:
        return err(502, f"费用闸门：连接智谱失败（{type(e).__name__}）。", "upstream_unreachable")
    p = 0
    try:
        p = int((resp.json().get("usage") or {}).get("prompt_tokens") or 0)
    except Exception:
        pass
    meta["task"] = "embedding"
    log_call(dept=dept, endpoint="embeddings", model=str(body.get("model", "")), stream=0,
             prompt_tokens=p, completion_tokens=0, cached_tokens=0, cost=p * PRICE_EMBED / 1_000_000,
             estimated=0, status=resp.status_code, latency_ms=int((time.monotonic() - t0) * 1000),
             note="" if resp.status_code == 200 else resp.text[:300], **meta)
    return Response(content=resp.content, status_code=resp.status_code, media_type="application/json")


# ----------------------------------------------------------------- 额度查询（给 Open WebUI 过滤器用）
@app.get("/api/spend")
async def api_spend(dept: str, chat_id: str = ""):
    if dept not in DEPTS:
        return err(404, f"未知部门代码 {dept}")
    info = spend_info(dept)
    if chat_id:
        info["chat_spent"] = round(get_chat_spent(chat_id), 4)
        info["chat_last_input"] = get_chat_last_input(chat_id)
        info["chat_hint"] = chat_hint(info["chat_spent"], info["chat_last_input"])
    info["chat_limit"] = CHAT_MAX_YUAN
    return info


@app.get("/health")
async def health():
    return {"status": "ok", "version": VERSION, "mock": MOCK}


# ----------------------------------------------------------------- 页面公共部分
CSS = """
*{box-sizing:border-box}body{margin:0;font:14px/1.6 -apple-system,"Microsoft YaHei","PingFang SC",sans-serif;color:#333;background:#F4F6F9}
header{background:#1F3B5C;color:#fff;padding:14px 28px;display:flex;align-items:center;gap:18px}
header h1{font-size:18px;margin:0}header a{color:#BFD6F0;text-decoration:none;margin-left:auto}
main{max-width:1180px;margin:22px auto;padding:0 16px}
.card{background:#fff;border-radius:8px;padding:18px 22px;margin-bottom:18px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
.card h2{font-size:16px;margin:0 0 12px;color:#1F3B5C}
table{border-collapse:collapse;width:100%}th,td{padding:9px 12px;border-bottom:1px solid #E6E9EE;text-align:left;white-space:nowrap}
th{background:#F7F9FC;color:#555;font-weight:600}td.num,th.num{text-align:right}
.ok{color:#3A9A5B;font-weight:700}.bad{color:#C0392B;font-weight:700}.warn{color:#D9822B;font-weight:700}
.bar{height:12px;background:#E9EDF2;border-radius:6px;overflow:hidden;min-width:160px}.bar i{display:block;height:100%}
.muted{color:#888;font-size:12px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
.kpi{background:#F7F9FC;border-radius:8px;padding:14px}.kpi b{font-size:24px;display:block}
input,button,select{font:inherit;padding:5px 9px;border:1px solid #C9D1DB;border-radius:5px}button{background:#2F6FB0;color:#fff;border:0;cursor:pointer}
.scroll{overflow-x:auto}table.wrap td{white-space:normal}table.wrap td:first-child{white-space:nowrap;width:150px;color:#555}
#internal-mark-bar{position:fixed;right:12px;bottom:10px;font-size:12px;color:#8a8a8a;background:rgba(255,255,255,.9);border:1px solid #e0e0e0;border-radius:4px;padding:2px 10px}
"""


MOCK_BANNER = ('<div style="background:#D9822B;color:#fff;padding:10px 28px;font-weight:700">'
               '【模拟模式】没有连接智谱：AI 回复是假的，费用是按字数估算的模拟值。只用于本机测试，NAS 正式环境的 GATE_MOCK 必须为 0。</div>')


def page(title, body, admin_link=True):
    link = '<a href="/admin">管理页 →</a>' if admin_link else '<a href="/">← 自检页</a>'
    return HTMLResponse(f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)}</title>
<style>{CSS}</style></head><body><header><h1>AI 平台 · {html.escape(title)}</h1>{link}</header>{MOCK_BANNER if MOCK else ""}
<main>{body}<p class="muted">cost-gate v{VERSION} · 启动于 {STARTED:%Y-%m-%d %H:%M} · 单价（元/百万 tokens）输入 {PRICE_IN} / 缓存命中 {PRICE_CACHE} / 输出 {PRICE_OUT}</p></main>
<div id="internal-mark-bar">内部文件，禁止外传</div></body></html>""")


def esc(v):
    return html.escape("" if v is None else str(v))


def budget_rows(month):
    rows = []
    for i, code in enumerate(DEPTS):
        s = spend_info(code, month)
        color = "#C0392B" if s["blocked"] else ("#D9822B" if s["warn"] else DEPT_COLORS[i % len(DEPT_COLORS)])
        width = min(s["pct"], 100)
        state = '<span class="bad">已拦截</span>' if s["blocked"] else ('<span class="warn">超过预警线</span>' if s["warn"] else '<span class="ok">正常</span>')
        rows.append((code, s, color, width, state))
    return rows


# ----------------------------------------------------------------- 自检页
async def check_upstream():
    if MOCK:
        return True, "模拟模式：不连接智谱", "本机测试用。正式环境请把 GATE_MOCK 设为 0。"
    t0 = time.monotonic()
    try:
        r = await _client.get(f"{UPSTREAM_BASE}/models", headers={"Authorization": f"Bearer {UPSTREAM_KEY or 'test'}"},
                              timeout=httpx.Timeout(10))
        ms = int((time.monotonic() - t0) * 1000)
        return True, f"通（HTTP {r.status_code}，{ms} ms）", ""
    except httpx.ConnectTimeout:
        return False, "不通：连接超时", "多半是公司防火墙没放行。请让 IT 放行 NAS 访问 open.bigmodel.cn 的 443 端口。"
    except httpx.ConnectError as e:
        msg = str(e)
        if "Name or service not known" in msg or "Temporary failure in name resolution" in msg or "nodename" in msg:
            return False, "不通：域名解析失败", "请检查 DSM 控制面板 → 网络 → DNS 服务器（可填 218.2.2.2 或公司路由器地址）。"
        if "CERTIFICATE" in msg.upper() or "SSL" in msg.upper():
            return False, "不通：证书校验失败", "公司可能有上网行为管理设备在解密 HTTPS，需要导入其根证书。"
        return False, f"不通：{esc(msg)[:120]}", "请截图发给平台管理员。"
    except Exception as e:
        return False, f"不通：{type(e).__name__}", "请截图发给平台管理员。"


@app.get("/", response_class=HTMLResponse)
async def selfcheck():
    ok_up, up_text, up_hint = await check_upstream()
    try:
        with _db_lock, db() as c:
            c.execute("CREATE TABLE IF NOT EXISTS _probe(x)")
            c.execute("DROP TABLE _probe")
            last = c.execute("SELECT ts FROM calls WHERE status=200 ORDER BY id DESC LIMIT 1").fetchone()
        db_ok, last_ok = True, (last["ts"] if last else "暂无")
    except Exception as e:
        db_ok, last_ok = False, f"{type(e).__name__}"
    init_items = []
    try:
        rep = json.load(open(_env("INIT_REPORT", "/init-report/init_report.json"), encoding="utf-8"))
        for it in rep.get("results", []):
            init_items.append((f"部门文件夹「AI{it['dept']}」", it["ok"], "写入正常" if it["ok"] else "终端无法写入",
                               "" if it["ok"] else esc("；".join(it["msgs"][-1:]))))
    except FileNotFoundError:
        init_items.append(("部门文件夹", False, "未找到初始化报告", "确认 ai-init-folders 容器已运行过（看它的日志）。"))
    except Exception as e:
        init_items.append(("部门文件夹", False, f"报告读取失败：{type(e).__name__}", ""))
    items = [
        ("费用闸门运行", True, f"正常（v{VERSION}）", ""),
        ("智谱 API 网络连通", ok_up, up_text, up_hint),
        ("智谱 API 密钥", True, "模拟模式：不需要密钥", "") if MOCK else ("智谱 API 密钥", bool(UPSTREAM_KEY), "已配置（有效性请在管理页点「测试密钥」）" if UPSTREAM_KEY else "未配置",
         "" if UPSTREAM_KEY else "在 .env 里填写 ZHIPU_API_KEY 后，重启 cost-gate 容器。"),
        ("部门密钥", bool(DEPTS), f"已配置 {len(DEPTS)} 个：" + "、".join(d["name"] for d in DEPTS.values()), ""),
        ("费用数据库", db_ok, "可读写" if db_ok else f"异常：{last_ok}", "" if db_ok else "检查 /data 挂载目录权限。"),
        ("用量保护", True, (f"单个对话上限 ¥{CHAT_MAX_YUAN:g}" if CHAT_MAX_YUAN else "单个对话不限") + "；" +
         (f"单次输入上限 {MAX_INPUT_TOKENS / 10000:g} 万 tokens" if MAX_INPUT_TOKENS else "单次输入不限") + "；" +
         (f"工具输出超过 {TOOL_OUTPUT_MAX_CHARS} 字截断" if TOOL_OUTPUT_MAX_CHARS else "工具输出不截断") + "；" +
         (f"长对话提示：输入 ≥ {CHAT_HINT_TOKENS / 10000:g} 万 tokens 或费用 ≥ {CHAT_HINT_PCT:g}%" if (CHAT_HINT_TOKENS or CHAT_HINT_PCT) else "不做长对话提示"),
         "在 .env 里用 CHAT_MAX_YUAN、MAX_INPUT_TOKENS、TOOL_OUTPUT_MAX_CHARS、CHAT_HINT_TOKENS、CHAT_HINT_PCT 调整"),
        ("管理员密码", bool(ADMIN_PASSWORD), "已设置" if ADMIN_PASSWORD else "未设置（管理页无法登录）",
         "" if ADMIN_PASSWORD else "在 .env 里填写 GATE_ADMIN_PASSWORD。"),
    ] + init_items
    trs = "".join(
        f"<tr><td>{esc(n)}</td><td class='{'ok' if good else 'bad'}'>{'✔' if good else '✘'} {t}</td><td class='muted'>{h}</td></tr>"
        for n, good, t, h in items)
    month = month_of()
    brs = "".join(
        f"<tr><td>{esc(s['name'])}</td><td class='num'>¥{s['spent']:.2f}</td><td class='num'>¥{s['limit']:.2f}</td>"
        f"<td><div class='bar'><i style='width:{w}%;background:{col}'></i></div></td><td class='num'>{s['pct']:.1f}%</td><td>{st}</td></tr>"
        for code, s, col, w, st in budget_rows(month))
    body = f"""
<div class="card"><h2>连通性自检</h2><table><tr><th>检查项</th><th>结果</th><th>处理办法</th></tr>{trs}</table>
<p class="muted">最近一次成功调用：{esc(last_ok)} · 刷新页面即重新检查</p></div>
<div class="card"><h2>本月（{month}）部门额度</h2><table><tr><th>部门</th><th class="num">已用</th><th class="num">上限</th><th>进度</th><th class="num">占比</th><th>状态</th></tr>{brs}</table></div>"""
    return page("自检", body)


# ----------------------------------------------------------------- 管理页
def auth_level(request: Request):
    """返回 'admin'（可改）/ 'viewer'（只读）/ None"""
    auth = request.headers.get("authorization", "")
    if not auth.lower().startswith("basic "):
        return None
    try:
        _user, _, pw = base64.b64decode(auth[6:]).decode("utf-8").partition(":")
    except Exception:
        return None
    if ADMIN_PASSWORD and hmac.compare_digest(pw, ADMIN_PASSWORD):
        return "admin"
    if VIEWER_PASSWORD and hmac.compare_digest(pw, VIEWER_PASSWORD):
        return "viewer"
    return None


def admin_ok(request: Request):
    return auth_level(request) == "admin"


def view_ok(request: Request):
    return auth_level(request) in ("admin", "viewer")


def need_auth():
    return Response("需要管理员密码", status_code=401, headers={"WWW-Authenticate": 'Basic realm="cost-gate admin", charset="UTF-8"'})


def svg_daily(month):
    with db() as c:
        rows = c.execute("SELECT day, dept, SUM(cost) s FROM calls WHERE month=? GROUP BY day, dept", (month,)).fetchall()
    y, m = map(int, month.split("-"))
    first = datetime(y, m, 1)
    nxt = datetime(y + (m == 12), m % 12 + 1, 1)
    days = [(first + timedelta(d)).strftime("%Y-%m-%d") for d in range((nxt - first).days)]
    val = {}
    for r in rows:
        val.setdefault(r["day"], {})[r["dept"]] = r["s"]
    mx = max([sum(v.values()) for v in val.values()] + [0.01])
    W, H, pad = 1100, 220, 40
    bw = (W - pad * 2) / len(days)
    parts = [f'<svg viewBox="0 0 {W} {H + 40}" width="100%" role="img" aria-label="按日费用">']
    for g in range(5):
        gy = H - (H - 20) * g / 4
        parts.append(f'<line x1="{pad}" x2="{W - pad}" y1="{gy:.1f}" y2="{gy:.1f}" stroke="#EEE"/>'
                     f'<text x="{pad - 6}" y="{gy + 4:.1f}" font-size="11" text-anchor="end" fill="#888">¥{mx * g / 4:.{3 if mx < 1 else 2}f}</text>')
    for i, d in enumerate(days):
        x = pad + i * bw
        base = H
        for j, code in enumerate(DEPTS):
            v = val.get(d, {}).get(code, 0)
            h = (H - 20) * v / mx
            if h > 0:
                parts.append(f'<rect x="{x + 2:.1f}" y="{base - h:.1f}" width="{bw - 4:.1f}" height="{h:.1f}" fill="{DEPT_COLORS[j % len(DEPT_COLORS)]}"><title>{d} {DEPTS[code]["name"]} ¥{v:.3f}</title></rect>')
                base -= h
        if i % 2 == 0:
            parts.append(f'<text x="{x + bw / 2:.1f}" y="{H + 16}" font-size="11" text-anchor="middle" fill="#888">{d[8:]}</text>')
    lx = pad
    for j, code in enumerate(DEPTS):
        parts.append(f'<rect x="{lx}" y="{H + 26}" width="12" height="10" fill="{DEPT_COLORS[j % len(DEPT_COLORS)]}"/>'
                     f'<text x="{lx + 16}" y="{H + 35}" font-size="12" fill="#555">{esc(DEPTS[code]["name"])}</text>')
        lx += 110
    parts.append("</svg>")
    return "".join(parts)


@app.get("/admin", response_class=HTMLResponse)
async def admin(request: Request, month: str = "", msg: str = ""):
    if not view_ok(request):
        return need_auth()
    is_admin = admin_ok(request)
    month = month or month_of()
    with db() as c:
        months = [r["month"] for r in c.execute("SELECT DISTINCT month FROM calls ORDER BY month DESC").fetchall()]
        tot = c.execute("SELECT COUNT(*) n, COALESCE(SUM(cost),0) s, COALESCE(SUM(prompt_tokens),0) p, COALESCE(SUM(completion_tokens),0) o, COALESCE(SUM(estimated),0) e FROM calls WHERE month=?", (month,)).fetchone()
        users = c.execute("""SELECT user_name, user_email, dept, COUNT(*) n, SUM(prompt_tokens+completion_tokens) t, SUM(cost) s
                             FROM calls WHERE month=? GROUP BY user_email, dept ORDER BY s DESC LIMIT 10""", (month,)).fetchall()
        tasks = c.execute("SELECT task, COUNT(*) n, SUM(cost) s FROM calls WHERE month=? GROUP BY task ORDER BY s DESC", (month,)).fetchall()
        chats = c.execute("""SELECT chat_id, MAX(user_name) u, dept, COUNT(*) n, SUM(prompt_tokens) p, SUM(cost) s, MIN(ts) t0, MAX(ts) t1
                             FROM calls WHERE month=? AND chat_id<>'' GROUP BY chat_id, dept ORDER BY s DESC LIMIT 10""", (month,)).fetchall()
        blocked = c.execute("""SELECT ts, user_name, dept, status, note FROM calls WHERE month=? AND status IN (402, 413)
                               ORDER BY id DESC LIMIT 20""", (month,)).fetchall()
        recent = c.execute("SELECT * FROM calls WHERE month=? ORDER BY id DESC LIMIT 50", (month,)).fetchall()
    if month_of() not in months:
        months.insert(0, month_of())
    opts = "".join(f'<option {"selected" if m == month else ""}>{m}</option>' for m in months)
    brs = "".join(
        f"<tr><td>{esc(s['name'])}</td><td class='num'>¥{s['spent']:.2f}</td>"
        + (f"<td><form method='post' action='/admin/budget' style='margin:0;display:flex;gap:6px'>"
           f"<input type='hidden' name='dept' value='{esc(code)}'><input name='limit' value='{s['limit']:.2f}' size='7'>"
           f"<button>保存</button></form></td>" if is_admin else f"<td class='num'>¥{s['limit']:.2f}</td>") +
        f"<td><div class='bar'><i style='width:{w}%;background:{col}'></i></div></td><td class='num'>{s['pct']:.1f}%</td><td>{st}</td></tr>"
        for code, s, col, w, st in budget_rows(month))
    urs = "".join(f"<tr><td>{esc(u['user_name'] or '（未知）')}</td><td>{esc(u['user_email'])}</td><td>{esc(DEPTS.get(u['dept'], {}).get('name', u['dept']))}</td>"
                  f"<td class='num'>{u['n']}</td><td class='num'>{(u['t'] or 0):,}</td><td class='num'>¥{(u['s'] or 0):.3f}</td></tr>" for u in users) or "<tr><td colspan=6 class='muted'>暂无</td></tr>"
    task_name = {"chat": "对话", "title_generation": "生成标题", "tags_generation": "生成标签", "follow_up_generation": "追问建议",
                 "embedding": "向量化", "query_generation": "检索词", "autocomplete_generation": "自动补全", "selfcheck": "密钥测试"}
    trs = "".join(f"<tr><td>{esc(task_name.get(t['task'], t['task']))}</td><td class='num'>{t['n']}</td><td class='num'>¥{(t['s'] or 0):.3f}</td></tr>" for t in tasks) or "<tr><td colspan=3 class='muted'>暂无</td></tr>"
    rrs = "".join(
        f"<tr><td>{esc(r['ts'][5:16].replace('T', ' '))}</td><td>{esc(r['user_name'])}</td><td>{esc(DEPTS.get(r['dept'], {}).get('name', r['dept']))}</td>"
        f"<td>{esc(task_name.get(r['task'], r['task']))}</td><td class='num'>{r['prompt_tokens']:,}</td><td class='num'>{r['cached_tokens']:,}</td>"
        f"<td class='num'>{r['completion_tokens']:,}</td><td class='num'>¥{r['cost']:.4f}{'*' if r['estimated'] else ''}</td>"
        f"<td class='num'>{r['latency_ms'] / 1000:.1f}s</td><td class='{'ok' if r['status'] == 200 else 'bad'}'>{r['status']}</td><td class='muted'>{esc(r['note'])[:60]}</td></tr>"
        for r in recent) or "<tr><td colspan=11 class='muted'>暂无</td></tr>"
    crs = "".join(
        f"<tr><td>{esc(r['u'] or '（未知）')}</td><td>{esc(DEPTS.get(r['dept'], {}).get('name', r['dept']))}</td>"
        f"<td class='muted'>{esc(r['chat_id'][:8])}…</td><td>{esc(r['t0'][5:16].replace('T', ' '))} ~ {esc(r['t1'][11:16])}</td>"
        f"<td class='num'>{r['n']}</td><td class='num'>{(r['p'] or 0) / max(r['n'], 1):,.0f}</td>"
        f"<td class='num {'bad' if CHAT_MAX_YUAN and (r['s'] or 0) >= CHAT_MAX_YUAN * 0.8 else ''}'>¥{(r['s'] or 0):.3f}</td></tr>"
        for r in chats) or "<tr><td colspan=7 class='muted'>暂无</td></tr>"
    brs2 = "".join(
        f"<tr><td>{esc(r['ts'][5:16].replace('T', ' '))}</td><td>{esc(r['user_name'])}</td>"
        f"<td>{esc(DEPTS.get(r['dept'], {}).get('name', r['dept']))}</td><td class='bad'>{r['status']}</td><td class='muted'>{esc(r['note'])}</td></tr>"
        for r in blocked) or "<tr><td colspan=5 class='muted'>本月没有被拦截的请求</td></tr>"
    notice = f"<div class='card' style='border-left:4px solid #3A9A5B'>{esc(msg)}</div>" if msg else ""
    if is_admin:
        admin_tools = f"""<form method="post" action="/admin/testkey" style="margin:0"><button type="submit" style="background:#D9822B">测试智谱密钥（约 ¥0.001，记入总裁办额度）</button></form></div>
<div class="card"><h2>平台初始化（Open WebUI）</h2>
<p class="muted">第一次部署时点一次；以后改了提示词或新增部门，再点一次即可（可重复执行，不会重复创建）。会自动：建部门组 → 部门终端只给本部门用 → 建部门助手 → 快捷指令 → 装「部门额度显示」过滤器 → 页面顶部横幅 → 保存配置基线（配置检查用）。</p>
<form method="post" action="/admin/setup" style="display:flex;gap:8px;flex-wrap:wrap">
<input name="email" placeholder="Open WebUI 管理员邮箱" size="28"><input name="password" type="password" placeholder="管理员密码" size="18"><button>一键初始化</button></form></div>
<div class="card"><h2>配置检查</h2>
<p class="muted">和最近一次「一键初始化」后的配置基线比较，看有没有人在 Open WebUI 界面上手工改过助手、提示词、快捷指令、权限、过滤器。升级 Open WebUI 版本后也点一次。</p>
<form method="post" action="/admin/drift" style="display:flex;gap:8px;flex-wrap:wrap">
<input name="email" placeholder="Open WebUI 管理员邮箱" size="28"><input name="password" type="password" placeholder="管理员密码" size="18"><button style="background:#8C8C8C">检查配置</button></form></div>
<div class="card"><h2>批量导入员工账号</h2>
<p class="muted">每行一个人：<b>姓名,登录名,部门,初始密码</b>。登录名可以用工号（登录时输入「工号@localhost」），部门填 {esc("、".join(d["name"] for c, d in DEPTS.items() ))}。重复导入不会重复建号。</p>
<form method="post" action="/admin/users"><textarea name="lines" rows="6" style="width:100%;font:13px monospace;padding:8px;border:1px solid #C9D1DB;border-radius:5px" placeholder="张三,1001,资材部,Abc12345&#10;李四,1002,财务部,Abc12345"></textarea>
<div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap"><input name="email" placeholder="Open WebUI 管理员邮箱" size="28"><input name="password" type="password" placeholder="管理员密码" size="18"><button>导入</button></div></form></div>
"""
    else:
        admin_tools = "<span class='muted'>只读模式：可查看和导出，不能修改额度</span></div>"
    body = f"""{notice}
<div class="card" style="display:flex;gap:12px;align-items:center;flex-wrap:wrap">
<form method="get" style="margin:0">月份 <select name="month" onchange="this.form.submit()">{opts}</select></form>
<a href="/admin/value?month={esc(month)}"><button type="button" style="background:#3A9A5B">月度价值 →</button></a>
<a href="/admin/export.csv?month={esc(month)}"><button type="button">导出本月明细（Excel 可打开）</button></a>
{admin_tools}<div class="card"><div class="grid">
<div class="kpi">本月调用次数<b>{tot['n']:,}</b></div><div class="kpi">本月费用<b>¥{tot['s']:.2f}</b></div>
<div class="kpi">输入 / 输出 tokens<b>{tot['p']:,} / {tot['o']:,}</b></div><div class="kpi">按字符估算的调用<b>{tot['e']}</b><span class="muted">带 * 号，金额为估算值</span></div></div></div>
<div class="card"><h2>部门额度{"（上限可直接修改，立即生效；填 0 = 暂停该部门）" if is_admin else ""}</h2><div class="scroll"><table><tr><th>部门</th><th class="num">已用</th><th>月上限（元）</th><th>进度</th><th class="num">占比</th><th>状态</th></tr>{brs}</table></div></div>
<div class="card"><h2>按日费用</h2>{svg_daily(month)}</div>
<div class="card" style="display:grid;grid-template-columns:2fr 1fr;gap:18px">
<div><h2>按人 Top 10</h2><div class="scroll"><table><tr><th>姓名</th><th>邮箱</th><th>部门</th><th class="num">次数</th><th class="num">tokens</th><th class="num">费用</th></tr>{urs}</table></div></div>
<div><h2>按用途</h2><table><tr><th>用途</th><th class="num">次数</th><th class="num">费用</th></tr>{trs}</table></div></div>
<div class="card"><h2>单个对话花费 Top 10</h2><p class="muted">单个对话上限 ¥{CHAT_MAX_YUAN:g}（红色 = 已超过 80%）；「平均输入」越大说明对话越长，建议提醒员工新建对话。</p>
<div class="scroll"><table><tr><th>用户</th><th>部门</th><th>对话</th><th>时间</th><th class="num">调用次数</th><th class="num">平均输入 tokens</th><th class="num">费用</th></tr>{crs}</table></div></div>
<div class="card"><h2>被拦截的请求（最近 20 条）</h2><p class="muted">402 = 部门额度或单对话上限用完；413 = 单次输入超过 {MAX_INPUT_TOKENS / 10000:g} 万 tokens</p>
<div class="scroll"><table><tr><th>时间</th><th>用户</th><th>部门</th><th>状态</th><th>原因</th></tr>{brs2}</table></div></div>
<div class="card"><h2>最近 50 次调用</h2><div class="scroll"><table><tr><th>时间</th><th>用户</th><th>部门</th><th>用途</th><th class="num">输入</th><th class="num">缓存命中</th><th class="num">输出</th><th class="num">费用</th><th class="num">耗时</th><th>状态</th><th>备注</th></tr>{rrs}</table></div></div>"""
    return page("费用管理", body, admin_link=False)


# ----------------------------------------------------------------- 月度价值页
def dept_folder(code):
    """部门代码 → 共享文件夹名（资材部 → 资材，总裁办 → 总裁办），与部门终端一致"""
    return DEPTS[code]["name"].rstrip("部")


def _week_month(name):
    """周次文件夹名（2026-W38）→ 该周周四所在月份；认不出返回 None"""
    m = re.match(r"^(\d{4})-?W(\d{1,2})", name)
    if not m:
        return None
    try:
        return datetime.fromisocalendar(int(m.group(1)), int(m.group(2)), 4).strftime("%Y-%m")
    except ValueError:
        return None


def _dir_month(path, name):
    m = re.match(r"^(\d{4})(\d{2})(\d{2})_", name)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    wm = _week_month(name)
    if wm:
        return wm
    try:
        return datetime.fromtimestamp(os.stat(path).st_mtime, CN_TZ).strftime("%Y-%m")
    except OSError:
        return None


def dept_assets(code):
    """扫部门文件夹（只看目录名、条目数）：{mounted, outputs:{月份:个数}, recipes, notes, weekly_set}"""
    root = os.path.join(DEPTS_ROOT, dept_folder(code))
    info = {"mounted": os.path.isdir(root), "outputs": {}, "recipes": 0, "notes": 0, "notes_chars": 0,
            "notes_review": None, "weekly_set": False}
    if not info["mounted"]:
        return info
    for sub in ("02_输出", "03_周报"):
        base = os.path.join(root, sub)
        try:
            names = os.listdir(base)
        except OSError:
            continue
        for n in names:
            p = os.path.join(base, n)
            if n.startswith(".") or not os.path.isdir(p):
                continue
            mo = _dir_month(p, n)
            if mo:
                info["outputs"][mo] = info["outputs"].get(mo, 0) + 1
    try:
        sk = os.path.join(root, "04_技能")
        info["recipes"] = sum(1 for n in os.listdir(sk) if not n.startswith(".") and os.path.isfile(os.path.join(sk, n, "验收.json")))
    except OSError:
        pass
    try:
        with open(os.path.join(root, "05_模板", "部门说明.md"), encoding="utf-8") as f:
            text = f.read()
        items = [ln.strip() for ln in text.splitlines() if ln.lstrip().startswith("- ")]
        info["notes"], info["notes_chars"] = len(items), sum(len(x) for x in items)
        m = re.search(r"^<!-- 最近复核：(\d{4}-\d{2}-\d{2})", text, re.M)
        info["notes_review"] = m.group(1) if m else None
    except OSError:
        pass
    info["weekly_set"] = os.path.isfile(os.path.join(root, "05_模板", "周报说明.md"))
    return info


def notes_status(a):
    """月度价值页「部门说明」一格：(显示文字, 是否标黄)"""
    if not a["mounted"]:
        return "—", False
    if not a["notes"]:
        return "<span class='muted'>还没有</span>", False
    days = None
    if a["notes_review"]:
        try:
            days = (now_cn().date() - datetime.strptime(a["notes_review"], "%Y-%m-%d").date()).days
        except ValueError:
            pass
    full = a["notes_chars"] > NOTES_MAX_CHARS * 0.8
    stale = a["notes"] >= 5 and (days is None or days > NOTES_REVIEW_DAYS)
    rv = f"{a['notes_review'][5:]} 复核" if a["notes_review"] else "未复核"
    txt = f"{a['notes']} 条 · {a['notes_chars']}/{NOTES_MAX_CHARS} 字 · {rv}"
    return txt, (full or stale)


def usage_by_dept(months):
    """按部门、月份汇总 gate.db：活跃人数、对话数、费用、被拦截次数、触顶对话数"""
    qs = ",".join("?" for _ in months)
    with db() as c:
        rows = c.execute(f"""SELECT dept, month,
            COUNT(DISTINCT CASE WHEN task='chat' AND status=200 AND user_email<>'' THEN user_email END) users,
            COUNT(DISTINCT CASE WHEN task='chat' AND status=200 AND chat_id<>'' THEN chat_id END) chats,
            COALESCE(SUM(cost),0) cost,
            SUM(CASE WHEN status IN (402,413) THEN 1 ELSE 0 END) blocked,
            COUNT(DISTINCT CASE WHEN status=402 AND note LIKE '这个对话已用%' THEN chat_id END) capped
            FROM calls WHERE month IN ({qs}) GROUP BY dept, month""", months).fetchall()
    out = {}
    for r in rows:
        out[(r["dept"], r["month"])] = dict(r)
    return out


def last_months(month, n=6):
    y, m = map(int, month.split("-"))
    res = []
    for _ in range(n):
        res.append(f"{y:04d}-{m:02d}")
        y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    return res[::-1]


def svg_stacked(labels, series, unit, title, height=200, decimals=0):
    """堆叠柱状图：series = [(名称, 颜色, [值…])]；每根柱顶标合计"""
    W, H, pad, top = 540, height, 44, 22
    tot = [sum(s[2][i] for s in series) for i in range(len(labels))]
    raw = max(tot + [0]) or 1
    step = 10 ** math.floor(math.log10(raw / 4))
    step = next(k * step for k in (1, 2, 2.5, 5, 10) if k * step * 4 >= raw)  # 刻度取整：1/2/2.5/5 × 10^n
    mx = step * 4
    tick = (lambda v: f"{v:,.{0 if step >= 1 else 2}f}")
    bw = (W - pad * 2) / max(len(labels), 1)
    fmt = (lambda v: f"{v:,.{decimals}f}")
    parts = [f'<svg viewBox="0 0 {W} {H + 56}" width="100%" role="img" aria-label="{esc(title)}">',
             f'<text x="{pad}" y="14" font-size="13" font-weight="700" fill="#1F3B5C">{esc(title)}</text>',
             f'<text x="{W - pad}" y="14" font-size="11" text-anchor="end" fill="#888">单位：{esc(unit)}</text>']
    for g in range(5):
        gy = H - (H - top - 10) * g / 4
        parts.append(f'<line x1="{pad}" x2="{W - pad}" y1="{gy:.1f}" y2="{gy:.1f}" stroke="#EEE"/>'
                     f'<text x="{pad - 6}" y="{gy + 4:.1f}" font-size="10" text-anchor="end" fill="#999">{tick(mx * g / 4)}</text>')
    for i, lab in enumerate(labels):
        x = pad + i * bw + bw * 0.2
        w = bw * 0.6
        base = H
        for name, color, vals in series:
            h = (H - top - 10) * vals[i] / mx
            if h > 0:
                parts.append(f'<rect x="{x:.1f}" y="{base - h:.1f}" width="{w:.1f}" height="{h:.1f}" fill="{color}">'
                             f'<title>{esc(lab)} {esc(name)} {fmt(vals[i])}</title></rect>')
                base -= h
        parts.append(f'<text x="{x + w / 2:.1f}" y="{base - 4:.1f}" font-size="11" text-anchor="middle" fill="#333">{fmt(tot[i])}</text>'
                     f'<text x="{x + w / 2:.1f}" y="{H + 15}" font-size="11" text-anchor="middle" fill="#666">{esc(lab)}</text>')
    lx = pad
    for name, color, _ in series:
        parts.append(f'<rect x="{lx}" y="{H + 30}" width="11" height="10" fill="{color}"/>'
                     f'<text x="{lx + 15}" y="{H + 39}" font-size="11" fill="#555">{esc(name)}</text>')
        lx += 18 + 13 * len(name) + 20
    parts.append("</svg>")
    return "".join(parts)


@app.get("/admin/value", response_class=HTMLResponse)
async def admin_value(request: Request, month: str = ""):
    if not view_ok(request):
        return need_auth()
    month = month if re.match(r"^\d{4}-\d{2}$", month or "") else month_of()
    months6 = last_months(month, 6)
    use = usage_by_dept(months6)
    assets = {code: dept_assets(code) for code in DEPTS}
    with db() as c:
        all_months = [r["month"] for r in c.execute("SELECT DISTINCT month FROM calls ORDER BY month DESC").fetchall()]
    if month_of() not in all_months:
        all_months.insert(0, month_of())
    if month not in all_months:
        all_months.append(month)
    opts = "".join(f'<option {"selected" if m == month else ""}>{m}</option>' for m in all_months)

    def u(code, mo, k):
        return (use.get((code, mo)) or {}).get(k) or 0

    def outs(code, mo):
        return assets[code]["outputs"].get(mo, 0)

    rows, T = [], {"users": 0, "chats": 0, "outs": 0, "cost": 0.0, "blocked": 0, "capped": 0, "recipes": 0, "notes": 0}
    for i, code in enumerate(DEPTS):
        a = assets[code]
        o = outs(code, month) if a["mounted"] else None
        r = {k: u(code, month, k) for k in ("users", "chats", "cost", "blocked", "capped")}
        for k in r:
            T[k] += r[k]
        T["outs"] += o or 0
        T["recipes"] += a["recipes"]
        T["notes"] += a["notes"]
        per_chat = f"¥{r['cost'] / r['chats']:.2f}" if r["chats"] else "—"
        ns = notes_status(a)
        if ns[1]:
            T["notes_warn"] = T.get("notes_warn", 0) + 1
        per_out = f"¥{r['cost'] / o:.2f}" if o else "—"
        o_txt = "<span class='muted'>未挂载</span>" if o is None else f"{o}"
        rows.append(
            f"<tr><td><i style='display:inline-block;width:10px;height:10px;background:{DEPT_COLORS[i % len(DEPT_COLORS)]};margin-right:6px'></i>{esc(DEPTS[code]['name'])}</td>"
            f"<td class='num'>{r['users']}</td><td class='num'>{r['chats']}</td><td class='num'><b>{o_txt}</b></td>"
            f"<td class='num'>¥{r['cost']:.2f}</td><td class='num'>{per_chat}</td><td class='num'>{per_out}</td>"
            f"<td class='num {'warn' if r['blocked'] else ''}'>{r['blocked']}</td><td class='num {'warn' if r['capped'] else ''}'>{r['capped']}</td>"
            f"<td class='num'>{a['recipes'] if a['mounted'] else '—'}</td><td class='num {'warn' if ns[1] else ''}'>{ns[0]}</td>"
            f"<td>{('<span class=ok>已设置</span>' if a['weekly_set'] else '<span class=muted>未设置</span>') if a['mounted'] else '—'}</td></tr>")
    per_out_all = f"¥{T['cost'] / T['outs']:.2f}" if T["outs"] else "—"
    labels = [m[2:].replace("-", "/") for m in months6]
    ser_out = [(DEPTS[c]["name"], DEPT_COLORS[i % len(DEPT_COLORS)], [outs(c, m) for m in months6]) for i, c in enumerate(DEPTS)]
    ser_cost = [(DEPTS[c]["name"], DEPT_COLORS[i % len(DEPT_COLORS)], [round(u(c, m, "cost"), 2) for m in months6]) for i, c in enumerate(DEPTS)]
    mounted = all(a["mounted"] for a in assets.values())
    warn_mount = "" if mounted else ("<div class='card' style='border-left:4px solid #D9822B'>部分部门文件夹没有挂载到费用闸门（docker-compose.yml 里 cost-gate 的 "
                                     "/depts 只读挂载），成果数、配方数显示为「未挂载」。同步新版 docker-compose.yml 后，项目停止 → 构建即可。</div>")
    if T.get("notes_warn"):
        warn_mount += (f"<div class='card' style='border-left:4px solid #D9822B'>有 {T['notes_warn']} 个部门的部门说明需要整理或复核（下表标黄）："
                       "请该部门员工在对话框输入 <b>/部门说明</b>（复核 / 整理部门说明）。部门说明是助手每次任务都读的口径，过长或过时会让结果变差、费用变高。</div>")
    body = f"""{warn_mount}
<div class="card" style="display:flex;gap:12px;align-items:center;flex-wrap:wrap">
<form method="get" style="margin:0">月份 <select name="month" onchange="this.form.submit()">{opts}</select></form>
<a href="/admin?month={esc(month)}"><button type="button" style="background:#8C8C8C">← 返回费用管理</button></a>
<span class="muted">回答「平台值不值」：有多少人在用、做出了多少成果、每个成果花多少钱、经验积累了多少</span></div>
<div class="card"><div class="grid">
<div class="kpi">活跃人数<b>{T['users']}</b><span class="muted">本月发过对话的员工</span></div>
<div class="kpi">完成成果<b>{T['outs']}</b><span class="muted">02_输出 + 03_周报 新增文件夹</span></div>
<div class="kpi">本月费用<b>¥{T['cost']:.2f}</b><span class="muted">每个成果平均 {per_out_all}</span></div>
<div class="kpi">经验积累<b>{T['recipes']} 配方 · {T['notes']} 条口径</b><span class="muted">部门配方 · 部门说明（累计）</span></div>
</div></div>
<div class="card" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:18px">
<div>{svg_stacked(labels, ser_out, "个", "近 6 个月完成成果数")}</div>
<div>{svg_stacked(labels, ser_cost, "元", "近 6 个月费用", decimals=2)}</div></div>
<div class="card"><h2>{esc(month)} 各部门明细</h2><div class="scroll"><table>
<tr><th>部门</th><th class="num">活跃人数</th><th class="num">对话数</th><th class="num">完成成果</th><th class="num">费用</th>
<th class="num">每个对话</th><th class="num">每个成果</th><th class="num">被拦截</th><th class="num">触顶对话</th>
<th class="num">部门配方</th><th class="num">部门说明</th><th>周报说明</th></tr>{''.join(rows)}
<tr style="font-weight:700;background:#F7F9FC"><td>合计</td><td class="num">{T['users']}</td><td class="num">{T['chats']}</td><td class="num">{T['outs']}</td>
<td class="num">¥{T['cost']:.2f}</td><td class="num">{f"¥{T['cost'] / T['chats']:.2f}" if T['chats'] else '—'}</td><td class="num">{per_out_all}</td>
<td class="num">{T['blocked']}</td><td class="num">{T['capped']}</td><td class="num">{T['recipes']}</td><td class="num">{T['notes']}</td><td></td></tr>
</table></div></div>
<div class="card"><h2>口径说明</h2><table class="wrap">
<tr><td>活跃人数</td><td>当月至少有一次成功对话的员工数（按登录账号去重；自动生成标题等后台调用不算）</td></tr>
<tr><td>对话数</td><td>当月有成功调用的对话个数（同一个对话来回多次只算 1 个）</td></tr>
<tr><td>完成成果</td><td>部门文件夹 02_输出 和 03_周报 下当月新增的子文件夹个数（助手每个任务建一个）。按文件夹名里的日期/周次归月，认不出时按修改时间；只看文件夹名，不读文件内容</td></tr>
<tr><td>每个对话 / 每个成果</td><td>当月费用 ÷ 对话数 / ÷ 完成成果数。成果增加而单价下降 = 平台越用越省</td></tr>
<tr><td>被拦截</td><td>因部门额度用完、单对话上限（¥{CHAT_MAX_YUAN:g}）或单次输入太长被拦下的请求次数</td></tr>
<tr><td>触顶对话</td><td>达到单对话上限的对话个数：多的话说明对话拖得太长或任务太大，适合拆分或存成部门配方</td></tr>
<tr><td>部门配方</td><td>当前累计：04_技能 里带「验收.json」的配方个数</td></tr>
<tr><td>部门说明</td><td>05_模板/部门说明.md 的条目数（每条都是用户确认过的口径）、字数 / 上限（{NOTES_MAX_CHARS} 字，每次任务都要读，越长越费钱）、最近复核日期。
<b style="color:#D9822B">标黄</b> = 超过上限的 80%，或 5 条以上且超过 {NOTES_REVIEW_DAYS} 天没复核：请该部门在对话框输入 /部门说明，助手会列出全部条目、指出重复或过时的，确认后记下复核日期</td></tr>
</table></div>"""
    return page("月度价值", body, admin_link=False)


# ----------------------------------------------------------------- 配置检查
@app.post("/admin/drift", response_class=HTMLResponse)
async def admin_drift(request: Request):
    if not admin_ok(request):
        return need_auth()
    from owui_setup import Setup
    f = await _form(request)
    st = Setup(OWUI_URL, f.get("email", ""), f.get("password", ""), DEPTS, EXEC_DEPT)
    try:
        when, diffs = await st.check_drift(BASELINE_DIR)
    except Exception as e:
        return _result_page("配置检查", [(False, str(e))])
    finally:
        await st.client.aclose()
    if not diffs:
        body = (f"<div class='card'><h2>配置检查</h2><p class='ok'>✔ 和基线（{esc(when)} 一键初始化后）完全一致，没有人在界面上改过本平台的配置。</p>"
                "<p><a href='/admin'><button>返回管理页</button></a></p></div>")
        return page("配置检查", body, admin_link=False)
    trs = "".join(
        f"<tr><td>{esc(p)}</td><td>{'<span class=muted>（无）</span>' if a is None else esc(a)}</td>"
        f"<td class='{'bad' if b is None else 'warn'}'>{'（已删除）' if b is None else esc(b)}</td></tr>" for p, a, b in diffs[:200])
    body = f"""<div class="card"><h2>配置检查：发现 {len(diffs)} 处和基线不同</h2>
<p class="muted">基线 = {esc(when)} 最近一次「一键初始化」后的配置。下面是之后在 Open WebUI 界面上被改动的地方。
<br>处理办法：改动是有意的 → 把它写进代码（owui_setup.py），本机测试后同步 NAS，再点一键初始化（会覆盖成代码里的设置并更新基线）；
不是有意的 → 直接点一键初始化恢复。</p>
<div class="scroll"><table><tr><th>位置</th><th>基线</th><th>现在</th></tr>{trs}</table></div>
<p><a href="/admin"><button>返回管理页</button></a></p></div>"""
    return page("配置检查", body, admin_link=False)


@app.post("/admin/budget")
async def admin_budget(request: Request):
    if not admin_ok(request):
        return need_auth()
    form = urllib.parse.parse_qs((await request.body()).decode("utf-8"))
    dept = (form.get("dept") or [""])[0]
    try:
        limit = round(float((form.get("limit") or [""])[0]), 2)
        assert limit >= 0 and dept in DEPTS
    except Exception:
        return RedirectResponse("/admin?msg=" + urllib.parse.quote("上限格式不对，请填数字"), status_code=303)
    with _db_lock, db() as c:
        c.execute("INSERT INTO budgets(dept, limit_yuan, updated_at) VALUES(?,?,?) ON CONFLICT(dept) DO UPDATE SET limit_yuan=excluded.limit_yuan, updated_at=excluded.updated_at",
                  (dept, limit, now_cn().isoformat(timespec="seconds")))
    return RedirectResponse("/admin?msg=" + urllib.parse.quote(f"已把 {DEPTS[dept]['name']} 的月上限改为 ¥{limit:.2f}"), status_code=303)


@app.post("/admin/testkey")
async def admin_testkey(request: Request):
    if not admin_ok(request):
        return need_auth()
    if MOCK:
        return RedirectResponse("/admin?msg=" + urllib.parse.quote("当前是模拟模式，不会连接智谱，无需测试密钥。"), status_code=303)
    dept = EXEC_DEPT if EXEC_DEPT in DEPTS else next(iter(DEPTS))
    body = {"model": MODELS[0], "messages": [{"role": "user", "content": "只回复：好"}], "max_tokens": 16, "stream": False}
    t0 = time.monotonic()
    try:
        r = await _client.post(f"{UPSTREAM_BASE}/chat/completions", headers=upstream_headers(), json=body, timeout=60)
        ut = usage_tuple((r.json() if r.status_code == 200 else {}).get("usage")) or (0, 0, 0)
        log_call(dept=dept, user_id="", user_name="管理员自检", user_email="", user_role="admin", chat_id="", task="selfcheck",
                 endpoint="chat", model=MODELS[0], stream=0, prompt_tokens=ut[0], completion_tokens=ut[1], cached_tokens=ut[2],
                 cost=calc_cost(*ut), estimated=0, status=r.status_code, latency_ms=int((time.monotonic() - t0) * 1000),
                 note="" if r.status_code == 200 else r.text[:200])
        if r.status_code == 200:
            msg = f"密钥有效：智谱正常应答（{int((time.monotonic() - t0) * 1000)} ms，消耗 {ut[0] + ut[1]} tokens）。"
        elif r.status_code == 401:
            msg = "密钥无效（HTTP 401）：请检查 .env 里的 ZHIPU_API_KEY。"
        elif r.status_code == 429:
            msg = f"智谱返回 429（额度不足或并发超限）：{r.text[:150]}"
        else:
            msg = f"智谱返回 HTTP {r.status_code}：{r.text[:150]}"
    except Exception as e:
        msg = f"连接智谱失败：{type(e).__name__}。请先看自检页的网络连通项。"
    return RedirectResponse("/admin?msg=" + urllib.parse.quote(msg), status_code=303)


@app.get("/admin/export.csv")
async def admin_export(request: Request, month: str = ""):
    if not view_ok(request):
        return need_auth()
    month = month or month_of()
    with db() as c:
        rows = c.execute("SELECT * FROM calls WHERE month=? ORDER BY id", (month,)).fetchall()
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["时间", "部门", "姓名", "邮箱", "角色", "对话ID", "用途", "接口", "模型", "流式", "输入tokens", "缓存命中tokens",
                "输出tokens", "费用(元)", "是否估算", "HTTP状态", "耗时(ms)", "备注"])
    for r in rows:
        w.writerow([r["ts"], DEPTS.get(r["dept"], {}).get("name", r["dept"]), r["user_name"], r["user_email"], r["user_role"],
                    r["chat_id"], r["task"], r["endpoint"], r["model"], r["stream"], r["prompt_tokens"], r["cached_tokens"],
                    r["completion_tokens"], f"{r['cost']:.6f}", "是" if r["estimated"] else "", r["status"], r["latency_ms"], r["note"]])
    w.writerow([])
    w.writerow(["内部文件，禁止外传"])
    data = "﻿" + buf.getvalue()
    fname = urllib.parse.quote(f"AI平台费用明细_{month}.csv")
    return Response(data.encode("utf-8"), media_type="text/csv; charset=utf-8",
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{fname}"})


async def _form(request):
    return {k: v[0] for k, v in urllib.parse.parse_qs((await request.body()).decode("utf-8"), keep_blank_values=True).items()}


def _result_page(title, log):
    lis = "".join(f"<tr><td class='{'ok' if good else 'bad'}'>{'✔' if good else '✘'}</td><td>{esc(m)}</td></tr>" for good, m in log)
    return page(title, f"<div class='card'><h2>{esc(title)}</h2><table>{lis}</table><p><a href='/admin'><button>返回管理页</button></a></p></div>", admin_link=False)


@app.post("/admin/setup", response_class=HTMLResponse)
async def admin_setup(request: Request):
    if not admin_ok(request):
        return need_auth()
    from owui_setup import Setup
    f = await _form(request)
    log = await Setup(OWUI_URL, f.get("email", ""), f.get("password", ""), DEPTS, EXEC_DEPT).run_all(GATE_INTERNAL_URL, list(DEPTS), BASELINE_DIR)
    return _result_page("平台初始化结果", log)


@app.post("/admin/users", response_class=HTMLResponse)
async def admin_users(request: Request):
    if not admin_ok(request):
        return need_auth()
    from owui_setup import Setup
    f = await _form(request)
    log = await Setup(OWUI_URL, f.get("email", ""), f.get("password", ""), DEPTS, EXEC_DEPT).import_users(f.get("lines", ""))
    return _result_page("账号导入结果", log)


if __name__ == "__main__":
    import sys
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    if not DEPTS:
        raise SystemExit("未配置 GATE_DEPTS，费用闸门无法启动。")
    if MOCK:
        print("=" * 60 + "\n【模拟模式】GATE_MOCK=1：不连接智谱，只用于本机测试\n" + "=" * 60, flush=True)
    uvicorn.run(app, host="0.0.0.0", port=int(_env("GATE_PORT", "8000")), log_level="info", access_log=False)
