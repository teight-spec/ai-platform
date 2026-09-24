# -*- coding: utf-8 -*-
"""
Open WebUI 一键初始化（由费用闸门管理页调用）
  - 建部门组、部门助手（绑定部门终端 + 系统提示词 + 只对本部门可见）
  - 给部门终端设置访问权限（只本部门可用）
  - 助手的开场建议按钮、模型参数（温度 0.3）
  - 斜杠快捷指令 /数据处理 /出图 /周报 /周报设置 /专题汇报（总裁办另有 /周报汇总）
  - 安装并启用全局过滤器「部门额度显示」
  - 页面顶部横幅（告知对话会被保存、总裁办和管理员可查看）
  - 初始化完成后保存一份「配置基线」，以后可在管理页检查有没有人在界面上手工改过配置
  - 批量导入账号并加入部门组
全部操作可重复执行：已存在就更新，不会重复创建。
"""
import hashlib
import json
import os
import time
from datetime import datetime

import httpx

HERE = os.path.dirname(os.path.abspath(__file__))
BASE_MODEL = os.getenv("GATE_MODELS", "glm-5.3-flash").split(",")[0].strip()
FILTER_ID = "dept_budget"

# 提示词的写法：固定内容在前、每次都变的内容（用户名、日期）放在最后。
# 上下文缓存一般按「开头相同的部分」命中（命中部分单价不到正常输入的 1/3，见 .env 的 PRICE_CACHED）；
# 变量放在开头，每人每天的开头都不一样，缓存就很难命中。效果看管理页「最近 50 次调用」的「缓存命中」列。
SYSTEM_PROMPT = """你是「{assistant}」，服务于公司{dept_name}，主要做三件事：数据处理、出图表、收集数据生成周报汇报 PPT。

## 工作环境
- 你可以使用终端（Linux + Python + Node + LibreOffice），终端没有外网，不能安装任何包。
- 本部门文件夹 ~/{folder}/ ，只在这里读写：01_原始数据、02_输出、03_周报、04_技能（部门技能和配方）、05_模板（部门说明、周报说明、往期样本）。对话里上传的文件会存进 01_原始数据。
- 公司技能总目录 /opt/company/skills/INDEX.md；公司工具包 /opt/company/kit/（dept.py 部门说明与配方、peek.py 看数据、charts 出图、tables 出 Excel、runio 可重跑脚本）。

## 每个任务先做
- 先运行 `python3 /opt/company/kit/dept.py start`：部门说明里已确认的口径直接照用，不再重复问；有同类部门配方就用配方重跑。
- 再按任务读技能（部门 04_技能 里有同类技能时以部门的为准）：
  - 数据处理 / 出图 / 汇总表 → /opt/company/skills/公司-数据处理与出图/SKILL.md
  - 周报 → /opt/company/skills/公司-周报PPT/SKILL.md（先按部门的 05_模板/周报说明.md 做；还没有就先了解部门现在怎么做周报，不要套用固定格式）
  - 记口径、存配方、用配方 → /opt/company/skills/公司-部门说明与配方/SKILL.md
  - 其他 Excel/Word/PPT/PDF 操作 → 先读 INDEX.md

## 必须遵守
1. 数字全部用代码计算，不许心算、估算或编造。回复里写明数据来源文件、行数、筛选条件。
2. 任何 Excel/CSV 先跑 `python3 /opt/company/kit/peek.py <文件>`，不要把整张表读进对话；终端只打印汇总结果（不超过 30 行）。
3. 字段含义不清楚、有多种理解时，先问用户，确认后再算。数据不足以支撑的结论标注「存疑」。
4. 输出放 ~/{folder}/02_输出/<YYYYMMDD>_<用户名>_<任务简称>/ （周报放 03_周报/<周次>/），计算代码存为同目录 script.py（用 runio 写成可重跑，结论数字写进 结果摘要.json）。临时文件用 `mktemp -d` 建自己的临时目录。
5. 交付前运行 `python3 /opt/company/stamp_internal.py <输出目录> -r` 加「内部文件，禁止外传」标注。PPT 一律转 PDF（/opt/company/office2pdf.py），只交 PDF。
6. 报表是给领导看的：先给结论（不超过 3 条，带数字），再给图表，明细放附件。
7. 任务结束时：本次新确认的口径列成「建议记入部门说明」，用户同意后用 dept.py note 写入；以后还会重复的任务，问用户要不要存成部门配方。
8. 用中文回复，简洁明确。

## 本次对话
- 当前用户：{{{{USER_NAME}}}}；今天：{{{{CURRENT_DATE}}}}"""


EXEC_PROMPT = """你是「{assistant}」，服务于公司总裁办，负责跨部门的数据汇总、出图和汇报材料。

## 工作环境
- 你可以使用终端（Linux + Python + Node + LibreOffice），终端没有外网，不能安装任何包。
- 各部门文件夹（只读）：{readonly_dirs}。每个部门下有 01_原始数据、02_输出、03_周报（每周一个 <周次> 子文件夹）、04_技能、05_模板（部门说明.md、周报说明.md 在这里）。
- 总裁办自己的文件夹（可写）：~/{folder}/ ，所有产出都放这里。
- 公司技能总目录 /opt/company/skills/INDEX.md；公司工具包 /opt/company/kit/（dept.py、peek.py、charts、tables、runio）。

## 每个任务先做
- 先运行 `python3 /opt/company/kit/dept.py start`（总裁办自己的部门说明和配方）。
- 用到某个部门的数据时，先读该部门的 05_模板/部门说明.md：那个部门的术语和口径以它为准。

## 必须遵守
1. 各部门文件夹只读，不要尝试修改、移动或删除里面的文件。
2. 汇总各部门周报时，读各部门 03_周报/<周次>/ 里已生成的周报（PDF 用 pdftotext 读，Excel 用 peek.py 看；有 结果摘要.json 优先读它），数字照抄并注明出处，不要重新从原始数据算一遍。
3. 数字全部用代码计算或照抄各部门周报，不许心算或估算；回复里写明引用了哪个部门的哪个文件。
4. 不同部门口径不一致时先指出差异再汇总；数据不足以支撑的结论标注「存疑」。
5. 任何 Excel/CSV 先跑 peek.py，不要把整张表读进对话。
6. 输出放 ~/{folder}/02_输出/<YYYYMMDD>_<任务简称>/，代码存为 script.py（用 runio 写成可重跑）；临时文件用 `mktemp -d`；交付前运行 `python3 /opt/company/stamp_internal.py <输出目录> -r`；PPT 一律转 PDF，只交 PDF。
7. 汇报材料先给结论（不超过 3 条，带数字），再给图表。任务结束时新确认的口径建议记入总裁办的部门说明，重复性任务问要不要存成配方。用中文回复，简洁明确。

## 本次对话
- 当前用户：{{{{USER_NAME}}}}；今天：{{{{CURRENT_DATE}}}}"""

# 页面顶部横幅（所有人可见，可关闭）：告知对话会被保存和查看；内容为空字符串 = 不显示横幅
BANNER_TEXT = ("本平台的对话和文件会保存在公司服务器上，总裁办和平台管理员可以查看全部对话。"
               "请不要发送身份证号、银行卡号等个人隐私。产出文件在本部门共享文件夹里。")
BANNER_ID = "ai-platform-notice-v1"   # 改了横幅内容时把 v1 改成 v2，已关闭横幅的人会重新看到一次

# 模型参数：数据类工作要稳定，温度调低
MODEL_PARAMS = {"temperature": 0.3}

# 开场建议（新对话页面上的按钮）
DEPT_STARTERS = [
    {"title": ["处理数据", "清洗、去重、汇总刚上传的 Excel"],
     "content": "请处理我刚上传的数据文件：先用 peek.py 看数据，告诉我每列你怎么理解、打算怎么清洗，等我确认后再执行。处理目标："},
    {"title": ["出图表", "按类别/日期汇总并出图 + Excel 汇总表"],
     "content": "请用刚上传的数据出图表：按【什么】汇总【哪个指标】，时间范围【】。先告诉我用什么图型，确认后再出图，并附 Excel 汇总表。"},
    {"title": ["出本周周报", "按部门的周报做法生成周报 PDF"],
     "content": "请按「公司-周报PPT」技能做本部门本周周报。先告诉我周次、日期范围和要用的材料，缺什么先说，确认后再做，最后给我 PDF 和 3 条结论。"},
    {"title": ["第一次设置周报", "先让助手了解我们现在的周报怎么做"],
     "content": "本部门第一次用 AI 做周报。我会把往期周报样本放进 05_模板，请按「公司-周报PPT」技能的情况 A 先了解我们现在的周报怎么做，逐项和我确认后写成周报说明。"},
]
EXEC_STARTERS = [
    {"title": ["汇总各部门周报", "读三部门本周周报，出一页总览"],
     "content": "请读取各部门 03_周报 里最新一周的周报，汇总成一页总览：每个部门 3 条要点 + 关键数字对比图，注明出处，只交 PDF。"},
    {"title": ["跨部门对比", "对比各部门关键指标的变化"],
     "content": "请对比各部门最近 4 周周报里的关键指标，找出变化最大的 3 项并出图，说明数据来自哪些文件。"},
]

# 斜杠快捷指令（在输入框打「/」弹出）。scope: dept=三个业务部门 + 总裁办，exec=只给总裁办
QUICK_COMMANDS = [
    {"command": "数据处理", "name": "数据处理（清洗/去重/汇总）", "scope": "all",
     "content": "请处理我刚上传的数据文件。先用 peek.py 看数据，告诉我每列你怎么理解、打算怎么清洗，等我确认后再执行。\n处理目标：【请填写，比如：去掉重复单号，按月汇总金额】"},
    {"command": "出图", "name": "出图表 + Excel 汇总表", "scope": "all",
     "content": "请用刚上传的数据出图表。\n汇总方式：按【什么】汇总【哪个指标】，时间范围【】\n先告诉我打算用什么图型，确认后出图，并附 Excel 汇总表。"},
    {"command": "周报", "name": "出本周周报（PDF）", "scope": "dept",
     "content": "请按「公司-周报PPT」技能做本部门周报（周次：【不填按本周】）。先告诉我周次、日期范围和要用的材料，缺什么先说，确认后再做，最后给我 PDF 和 3 条结论。"},
    {"command": "周报设置", "name": "第一次设置周报（先了解现在的做法）", "scope": "dept",
     "content": "本部门第一次用 AI 做周报。往期周报样本我放在 05_模板 里了（文件：【文件名】），请按「公司-周报PPT」技能的情况 A 先了解我们现在的周报怎么做，逐项和我确认后写成周报说明。"},
    {"command": "专题汇报", "name": "专题汇报 PPT（PDF）", "scope": "all",
     "content": "请把【主题/数据文件】做成汇报 PPT，约【页数】页，汇报对象【谁】。先给我提纲确认，再出 PPT，只交 PDF。"},
    {"command": "周报汇总", "name": "汇总各部门本周周报", "scope": "exec",
     "content": "请读取各部门 03_周报 里最新一周（周次：【不填按最新】）的周报，汇总成总览：每个部门 3 条要点 + 关键数字对比图，注明出处，只交 PDF。"},
]


def _short(v, n=60):
    t = v if isinstance(v, str) else json.dumps(v, ensure_ascii=False)
    t = t.replace("\n", "⏎")
    return t if len(t) <= n else t[: n - 1] + "…"


def _diff(a, b, path, out):
    """递归比较两份快照，结果追加到 out：(位置, 基线值, 现在值)；基线值/现在值为 None 表示新增/删除"""
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b)):
            p = f"{path} › {k}" if path else k
            if k not in b:
                out.append((p, _short(a[k]), None))
            elif k not in a:
                out.append((p, None, _short(b[k])))
            else:
                _diff(a[k], b[k], p, out)
    elif a != b:
        if isinstance(a, str) and isinstance(b, str) and (len(a) > 60 or len(b) > 60):
            out.append((path, f"（{len(a)} 字）{_short(a, 40)}", f"（{len(b)} 字）{_short(b, 40)}"))
        else:
            out.append((path, _short(a), _short(b)))


class Setup:
    def __init__(self, owui_url, email, password, depts, exec_code):
        self.url = owui_url.rstrip("/")
        self.email, self.password = email, password
        self.depts = dict(depts)
        self.exec_code = exec_code
        self.log = []
        self.client = httpx.AsyncClient(timeout=30, trust_env=False)
        self.h = {}

    def ok(self, msg):
        self.log.append((True, msg))

    def bad(self, msg):
        self.log.append((False, msg))

    async def req(self, method, path, **kw):
        r = await self.client.request(method, self.url + path, headers=self.h, **kw)
        if r.status_code >= 400:
            raise RuntimeError(f"{method} {path} → HTTP {r.status_code}: {r.text[:200]}")
        return r.json() if r.content else None

    async def signin(self):
        r = await self.client.post(self.url + "/api/v1/auths/signin", json={"email": self.email, "password": self.password})
        if r.status_code != 200:
            raise RuntimeError(f"登录 Open WebUI 失败（HTTP {r.status_code}）：请确认管理员邮箱和密码。")
        data = r.json()
        if data.get("role") != "admin":
            raise RuntimeError("这个账号不是管理员。请用第一个注册的账号（管理员）。")
        self.h = {"Authorization": f"Bearer {data['token']}"}
        self.ok(f"已用管理员 {data.get('name')} 登录 Open WebUI")

    async def groups(self):
        return {g["name"]: g for g in (await self.req("GET", "/api/v1/groups/") or [])}

    async def ensure_groups(self):
        existing = await self.groups()
        ids = {}
        for code, d in self.depts.items():
            if d["name"] in existing:
                ids[code] = existing[d["name"]]["id"]
                self.ok(f"部门组「{d['name']}」已存在")
            else:
                g = await self.req("POST", "/api/v1/groups/create", json={"name": d["name"], "description": f"{d['name']}成员（AI 平台）"})
                ids[code] = g["id"]
                self.ok(f"已创建部门组「{d['name']}」")
        return ids

    async def ensure_terminal_grants(self, gids):
        cfg = await self.req("GET", "/api/v1/configs/terminal_servers")
        conns = cfg.get("TERMINAL_SERVER_CONNECTIONS", [])
        found = set()
        for c in conns:
            code = (c.get("id") or "").replace("term-", "")
            if code in gids:
                c.setdefault("config", {})
                c["config"] = c["config"] or {}
                c["config"]["access_grants"] = [{"principal_type": "group", "principal_id": gids[code], "permission": "read"}]
                c["config"]["chat_uploads"] = "filesystem"  # 对话里上传的文件直接存进部门终端的当前目录
                found.add(code)
        await self.req("POST", "/api/v1/configs/terminal_servers", json={"TERMINAL_SERVER_CONNECTIONS": conns})
        for code in gids:
            name = self.depts[code]["name"]
            if code in found:
                self.ok(f"终端 term-{code} 已设为只有「{name}」可用，对话上传的文件直接存入部门文件夹")
            else:
                self.bad(f"没找到终端连接 term-{code}，请检查 docker-compose 里的 TERMINAL_SERVER_CONNECTIONS")

    async def ensure_base_models(self, gids):
        """部门助手要能用，底层的「部门连接模型」也必须对该部门授权；设为隐藏，员工列表里看不到它"""
        for code, gid in gids.items():
            d = self.depts[code]
            mid = f"{code}.{BASE_MODEL}"
            form = {"id": mid, "name": f"{d['name']}·基础模型", "base_model_id": None, "is_active": True,
                    "meta": {"description": "部门助手的底层连接，已隐藏", "hidden": True}, "params": {},
                    "access_grants": [{"principal_type": "group", "principal_id": gid, "permission": "read"}]}
            r = await self.client.get(self.url + "/api/v1/models/model", params={"id": mid}, headers=self.h)
            if r.status_code == 200 and r.content and r.json():
                await self.req("POST", "/api/v1/models/model/update", json=form)
                await self.req("POST", "/api/v1/models/model/access/update", json={"id": mid, "access_grants": form["access_grants"]})
            else:
                await self.req("POST", "/api/v1/models/create", json=form)
            self.ok(f"底层连接 {mid} 已授权给「{d['name']}」并隐藏")

    async def ensure_models(self, gids):
        for code, gid in gids.items():
            d = self.depts[code]
            folder = d["name"].rstrip("部")
            assistant = f"{folder}助手"
            mid = f"assistant-{code}"
            form = {
                "id": mid, "name": assistant, "base_model_id": f"{code}.{BASE_MODEL}", "is_active": True,
                "meta": {"description": ("总裁办专用：跨部门汇总分析（各部门数据只读）" if code == self.exec_code else f"{d['name']}专用：数据清洗、统计、出图、报表与周报"),
                         "terminalId": f"term-{code}",
                         "suggestion_prompts": EXEC_STARTERS if code == self.exec_code else DEPT_STARTERS,
                         "capabilities": {"vision": False, "file_upload": False, "web_search": False, "image_generation": False,
                                          "code_interpreter": False, "citations": False, "usage": True, "terminal": True, "builtin_tools": False,
                                          "status_updates": True}},
                "params": {"system": (EXEC_PROMPT.format(assistant=assistant, folder=folder,
                                                         readonly_dirs="、".join(f"~/各部门/{x['name'].rstrip('部')}/" for c2, x in self.depts.items() if c2 != self.exec_code))
                                      if code == self.exec_code else
                                      SYSTEM_PROMPT.format(assistant=assistant, dept_name=d["name"], folder=folder)),
                           "function_calling": "native", **MODEL_PARAMS},
                "access_grants": [{"principal_type": "group", "principal_id": gid, "permission": "read"}],
            }
            r = await self.client.get(self.url + "/api/v1/models/model", params={"id": mid}, headers=self.h)
            if r.status_code == 200 and r.content and r.json():
                await self.req("POST", "/api/v1/models/model/update", json=form)
                await self.req("POST", "/api/v1/models/model/access/update", json={"id": mid, "access_grants": form["access_grants"]})
                self.ok(f"已更新「{assistant}」（绑定 term-{code}，仅{d['name']}可见）")
            else:
                await self.req("POST", "/api/v1/models/create", json=form)
                self.ok(f"已创建「{assistant}」（绑定 term-{code}，仅{d['name']}可见）")

    async def ensure_prompts(self, gids):
        """斜杠快捷指令：按部门组授权；已存在就更新（可重复执行）"""
        existing = {}
        try:
            for p in await self.req("GET", "/api/v1/prompts/") or []:
                existing[(p.get("command") or "").lstrip("/")] = p
        except Exception as e:
            self.bad(f"读取快捷指令失败，已跳过：{e}")
            return
        exec_gid = gids.get(self.exec_code)
        for qc in QUICK_COMMANDS:
            if qc["scope"] == "exec":
                targets = [exec_gid]
            else:
                targets = list(gids.values())
            grants = [{"principal_type": "group", "principal_id": g, "permission": "read"} for g in targets if g]
            form = {"command": qc["command"], "name": qc["name"], "content": qc["content"],
                    "tags": ["AI平台"], "access_grants": grants, "commit_message": "一键初始化"}
            try:
                old = existing.get(qc["command"])
                if old:
                    await self.req("POST", f"/api/v1/prompts/id/{old['id']}/update", json=form)
                    await self.req("POST", f"/api/v1/prompts/id/{old['id']}/access/update", json={"access_grants": grants})
                else:
                    await self.req("POST", "/api/v1/prompts/create", json=form)
                who = "总裁办" if qc["scope"] == "exec" else "所有部门"
                self.ok(f"快捷指令 /{qc['command']}（{qc['name']}）已{'更新' if old else '创建'}，{who}可用")
            except Exception as e:
                self.bad(f"快捷指令 /{qc['command']} 设置失败：{e}")

    async def ensure_filter(self, gate_internal_url, all_codes):
        content = open(os.path.join(HERE, "dept_budget_filter.py"), encoding="utf-8").read()
        form = {"id": FILTER_ID, "name": "部门额度显示", "content": content,
                "meta": {"description": "对话前显示本部门本月已用金额，额度用完时提示"}}
        r = await self.client.get(self.url + f"/api/v1/functions/id/{FILTER_ID}", headers=self.h)
        if r.status_code == 200 and r.content and r.json():
            fn = await self.req("POST", f"/api/v1/functions/id/{FILTER_ID}/update", json=form)
        else:
            fn = await self.req("POST", "/api/v1/functions/create", json=form)
        fn = await self.req("GET", f"/api/v1/functions/id/{FILTER_ID}")
        if not fn.get("is_active"):
            await self.req("POST", f"/api/v1/functions/id/{FILTER_ID}/toggle")
        if not fn.get("is_global"):
            await self.req("POST", f"/api/v1/functions/id/{FILTER_ID}/toggle/global")
        await self.req("POST", f"/api/v1/functions/id/{FILTER_ID}/valves/update",
                       json={"gate_url": gate_internal_url, "dept_codes": ",".join(all_codes), "priority": 0})
        fn = await self.req("GET", f"/api/v1/functions/id/{FILTER_ID}")
        if fn.get("is_active") and fn.get("is_global"):
            self.ok("过滤器「部门额度显示」已安装，状态：启用 + 全局")
        else:
            self.bad(f"过滤器状态异常：启用={fn.get('is_active')} 全局={fn.get('is_global')}")

    async def ensure_banner(self):
        """页面顶部横幅：告知员工对话会保存、可被查看。BANNER_TEXT 为空 = 去掉本平台的横幅（别的横幅不动）"""
        cur = await self.req("GET", "/api/v1/configs/banners") or []
        keep = [b for b in cur if not str(b.get("id", "")).startswith("ai-platform-notice")]
        if BANNER_TEXT:
            old = next((b for b in cur if b.get("id") == BANNER_ID), None)
            keep.append({"id": BANNER_ID, "type": "info", "title": "", "content": BANNER_TEXT, "dismissible": True,
                         "timestamp": (old or {}).get("timestamp") or int(time.time())})
        await self.req("POST", "/api/v1/configs/banners", json={"banners": keep})
        self.ok("页面顶部横幅已设置：" + BANNER_TEXT if BANNER_TEXT else "已去掉本平台的页面横幅")

    # ------------------------------------------------------------ 配置快照（检查有没有人在界面上手工改过）
    async def snapshot(self):
        """把和本平台有关的 Open WebUI 配置整理成一份可比较的 JSON（去掉时间戳、密钥、成员名单这类会正常变化的内容）"""
        groups = await self.req("GET", "/api/v1/groups/") or []
        gname = {g["id"]: g["name"] for g in groups}

        def grants(gs):
            return sorted(f"{gname.get(g.get('principal_id'), g.get('principal_type', '') + ':' + str(g.get('principal_id')))}"
                          f"·{g.get('permission')}" for g in (gs or []))

        snap = {"部门组": {}, "助手与模型": {}, "快捷指令": {}, "过滤器": {}, "终端连接": {}, "横幅": [], "默认权限": {}}
        for g in groups:
            snap["部门组"][g["name"]] = {"说明": g.get("description", ""), "权限": g.get("permissions") or {}}
        for m in await self.req("GET", "/api/v1/models/export") or []:
            meta = dict(m.get("meta") or {})
            meta.pop("profile_image_url", None)
            snap["助手与模型"][m["id"]] = {"名称": m.get("name"), "底层模型": m.get("base_model_id"),
                                          "启用": m.get("is_active"), "参数": m.get("params") or {}, "设置": meta,
                                          "可见范围": grants(m.get("access_grants"))}
        for p in await self.req("GET", "/api/v1/prompts/") or []:
            snap["快捷指令"]["/" + (p.get("command") or "").lstrip("/")] = {
                "名称": p.get("name"), "内容": p.get("content"), "可见范围": grants(p.get("access_grants")),
                "启用": p.get("is_active", True)}
        for f in await self.req("GET", "/api/v1/functions/") or []:
            full = await self.req("GET", f"/api/v1/functions/id/{f['id']}") or {}
            try:
                valves = await self.req("GET", f"/api/v1/functions/id/{f['id']}/valves") or {}
            except Exception:
                valves = {}
            snap["过滤器"][f["id"]] = {"名称": f.get("name"), "启用": full.get("is_active"), "全局": full.get("is_global"),
                                     "设置": valves,
                                     "代码指纹": hashlib.sha256((full.get("content") or "").encode()).hexdigest()[:16]}
        cfg = await self.req("GET", "/api/v1/configs/terminal_servers") or {}
        for c in cfg.get("TERMINAL_SERVER_CONNECTIONS", []):
            c = {k: v for k, v in c.items() if k not in ("key",)}
            conf = dict(c.get("config") or {})
            conf["access_grants"] = grants(conf.get("access_grants"))
            c["config"] = conf
            snap["终端连接"][c.get("id") or c.get("url")] = c
        snap["横幅"] = [{k: b.get(k) for k in ("id", "type", "content", "dismissible")}
                      for b in (await self.req("GET", "/api/v1/configs/banners") or [])]
        try:
            snap["默认权限"] = await self.req("GET", "/api/v1/users/default/permissions") or {}
        except Exception as e:
            snap["默认权限"] = {"读取失败": str(e)[:80]}
        return snap

    async def save_baseline(self, folder):
        snap = await self.snapshot()
        os.makedirs(folder, exist_ok=True)
        doc = {"_说明": "一键初始化后的 Open WebUI 配置基线",
               "_时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "配置": snap}
        with open(os.path.join(folder, "基线.json"), "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        with open(os.path.join(folder, f"基线_{datetime.now():%Y%m%d_%H%M%S}.json"), "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        self.ok(f"已保存配置基线（{doc['_时间']}）：以后可在管理页「配置检查」发现界面上的手工改动")

    async def check_drift(self, folder):
        """和基线比较，返回 (基线时间, 差异列表)。差异 = [(位置, 基线值, 现在值)]"""
        p = os.path.join(folder, "基线.json")
        if not os.path.isfile(p):
            raise RuntimeError("还没有配置基线：请先在管理页点一次「一键初始化」")
        base = json.load(open(p, encoding="utf-8"))
        await self.signin()
        now = await self.snapshot()
        diffs = []
        _diff(base["配置"], now, "", diffs)
        return base.get("_时间", ""), diffs

    async def run_all(self, gate_internal_url, all_codes, baseline_dir=None):
        try:
            await self.signin()
            gids = await self.ensure_groups()
            await self.ensure_terminal_grants(gids)
            await self.ensure_base_models(gids)
            await self.ensure_models(gids)
            await self.ensure_prompts(gids)
            await self.ensure_filter(gate_internal_url, all_codes)
            await self.ensure_banner()
            if baseline_dir:
                await self.save_baseline(baseline_dir)
            self.ok("初始化完成。下一步：在下方批量导入员工账号。")
        except Exception as e:
            self.bad(str(e))
        finally:
            await self.client.aclose()
        return self.log

    async def import_users(self, text):
        try:
            await self.signin()
            gids = await self.ensure_groups()
            name_to_code = {}
            for code, d in self.depts.items():
                name_to_code[d["name"]] = code
                name_to_code[d["name"].rstrip("部")] = code
            n_ok = 0
            for ln, line in enumerate(text.splitlines(), 1):
                line = line.strip().replace("，", ",")
                if not line or line.startswith("#"):
                    continue
                parts = [p.strip() for p in line.split(",")]
                if len(parts) < 4:
                    self.bad(f"第 {ln} 行格式不对（应为：姓名,登录名,部门,初始密码）：{line}")
                    continue
                name, login, dept, pw = parts[:4]
                code = name_to_code.get(dept)
                if not code:
                    self.bad(f"第 {ln} 行部门「{dept}」不存在（可用：{'、'.join(d['name'] for d in self.depts.values())}）")
                    continue
                email = login if "@" in login else f"{login}@localhost"
                role = "admin" if code == self.exec_code else "user"
                r = await self.client.post(self.url + "/api/v1/auths/add", headers=self.h,
                                           json={"name": name, "email": email.lower(), "password": pw, "role": role})
                if r.status_code == 200:
                    uid = r.json()["id"]
                    note = "新建"
                else:
                    s = await self.req("GET", "/api/v1/users/search", params={"query": email.lower()})
                    users = s.get("users", []) if isinstance(s, dict) else []
                    match = [u for u in users if (u.get("email") or "").lower() == email.lower()]
                    if not match:
                        self.bad(f"第 {ln} 行 {name}：创建失败（{r.text[:120]}）")
                        continue
                    uid = match[0]["id"]
                    note = "已存在"
                await self.req("POST", f"/api/v1/groups/id/{gids[code]}/users/add", json={"user_ids": [uid]})
                self.ok(f"{name}（登录名 {email}）{note}，已加入「{self.depts[code]['name']}」"
                        + ("，角色：管理员（可查看全部对话和各部门文件）" if code == self.exec_code and note == "新建" else ""))
                n_ok += 1
            self.ok(f"共处理成功 {n_ok} 人。员工用「登录名」+「初始密码」登录后，请提醒他们在 设置 → 账号 里改密码。")
        except Exception as e:
            self.bad(str(e))
        finally:
            await self.client.aclose()
        return self.log
