"""
title: 部门额度显示
author: AI 平台管理员
version: 1.1.0
description: 每次对话前，在回复上方显示「本部门本月已用 ¥x / ¥y · 本对话 ¥a / ¥b」；额度用完时直接提示，不再调用模型。
"""
from pydantic import BaseModel, Field
import httpx


class Filter:
    class Valves(BaseModel):
        gate_url: str = Field(default="http://cost-gate:8000", description="费用闸门的内部地址")
        dept_codes: str = Field(default="zc,cw,yx,gl", description="部门代码（= 模型连接前缀），逗号分隔")
        priority: int = Field(default=0, description="过滤器执行顺序")

    def __init__(self):
        self.valves = self.Valves()

    def _dept(self, model: dict):
        codes = {c.strip() for c in self.valves.dept_codes.split(",") if c.strip()}
        info = (model or {}).get("info") or {}
        for mid in (info.get("base_model_id") or "", (model or {}).get("id") or ""):
            if "." in mid and mid.split(".", 1)[0] in codes:
                return mid.split(".", 1)[0]
        return None

    async def inlet(self, body: dict, __user__: dict = None, __model__: dict = None, __event_emitter__=None,
                    __metadata__: dict = None) -> dict:
        dept = self._dept(__model__)
        if not dept:
            return body
        chat_id = (__metadata__ or {}).get("chat_id") or ""
        try:
            async with httpx.AsyncClient(timeout=3, trust_env=False) as client:
                r = await client.get(f"{self.valves.gate_url.rstrip('/')}/api/spend", params={"dept": dept, "chat_id": chat_id})
                s = r.json()
        except Exception:
            return body  # 查不到额度时不影响正常使用，费用闸门仍会兜底拦截
        if s.get("blocked"):
            raise Exception(f"【{s['name']}】本月 AI 额度已用完（已用 ¥{s['spent']:.2f} / 上限 ¥{s['limit']:.2f}），"
                            f"请联系平台管理员，下月 1 日自动恢复。")
        text = f"{s['name']} · 本月已用 ¥{s['spent']:.2f} / ¥{s['limit']:.2f}（{s['pct']:.0f}%）"
        if s.get("chat_limit") and chat_id and not chat_id.startswith("local:"):
            text += f" · 本对话 ¥{s.get('chat_spent', 0):.2f} / ¥{s['chat_limit']:g}"
            if s.get("chat_spent", 0) >= s["chat_limit"]:
                raise Exception(f"这个对话已用 ¥{s['chat_spent']:.2f}，达到单个对话上限 ¥{s['chat_limit']:g}。"
                                f"请点「新对话」重新开始（需要的结论可以复制过去，文件都在部门文件夹里）。")
        if s.get("warn"):
            text += "　已超过预警线，请节约使用"
        if __event_emitter__:
            await __event_emitter__({"type": "status", "data": {"description": text, "done": True, "hidden": False}})
        return body
