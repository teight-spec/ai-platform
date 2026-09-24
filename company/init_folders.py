# -*- coding: utf-8 -*-
"""
部门文件夹初始化（每次启动平台时自动运行一次，可重复执行）
  1. 在每个部门共享文件夹里建好 01_原始数据 / 02_输出 / 03_周报 / 04_技能 / 05_模板
  2. 通用技能已改为公司层（/opt/company/skills/公司-*，随平台更新）：
     清理早期自动放进 04_技能 的旧模板副本（内容没被部门改过才删；改过的保留并提示）
  3. 用终端账号（群晖终端服务账号，UID 1000）实际写一次文件，检查权限；结果写到 init_report.json，费用闸门自检页会显示
"""
import hashlib
import json
import os
import shutil
from datetime import datetime

BASE = "/depts"
REPORT = "/report/init_report.json"
# 旧版（v1.0~1.1）自动复制到 04_技能 的模板：SKILL.md 的 sha256，一致=部门没改过，可以安全清理
OLD_TEMPLATES = {
    "周报生成": "9dfb65773ef735697815860323eea7ce107aba58b5901d68883a61ab39f5520b",
    "数据清洗与出图": "13dd4410441503a7f73a2bc9d44e294a60433bdf9d00e072c8668685dcc702ab",
}
SUBDIRS = ["01_原始数据", "02_输出", "03_周报", "04_技能", "05_模板"]
UID = int(os.getenv("TERMINAL_UID", "1000"))
GID = int(os.getenv("TERMINAL_GID", "100"))

results = []
for name in sorted(os.listdir(BASE)):
    root = os.path.join(BASE, name)
    item = {"dept": name, "ok": True, "msgs": []}
    try:
        for s in SUBDIRS:
            os.makedirs(os.path.join(root, s), exist_ok=True)
        for sk, digest in OLD_TEMPLATES.items():
            d = os.path.join(root, "04_技能", sk)
            f = os.path.join(d, "SKILL.md")
            if not os.path.isfile(f):
                continue
            with open(f, "rb") as fh:
                same = hashlib.sha256(fh.read()).hexdigest() == digest
            if same and sorted(os.listdir(d)) == ["SKILL.md"]:
                shutil.rmtree(d)
                item["msgs"].append(f"已清理旧模板 04_技能/{sk}（改用公司技能）")
            else:
                item["msgs"].append(f"04_技能/{sk} 被部门改过，已保留；它会优先于公司技能，请管理员对照新版公司技能决定是否删除")
        for s in SUBDIRS:  # 尽量放开权限；群晖 ACL 文件夹可能不允许，失败不要紧，以写入测试为准
            try:
                os.chmod(os.path.join(root, s), 0o777)
            except Exception:
                pass
        pid = os.fork()
        if pid == 0:
            code = 0
            try:
                if UID != 0:
                    os.setgroups([])
                    os.setgid(GID)
                    os.setuid(UID)
                p = os.path.join(root, "02_输出", ".写入测试.tmp")
                with open(p, "w") as f:
                    f.write("ok")
                os.remove(p)
            except Exception:
                code = 3
            os._exit(code)
        _, status = os.waitpid(pid, 0)
        if os.WEXITSTATUS(status) != 0:
            item["ok"] = False
            item["msgs"].append("终端账号（UID %d）无法写入。请检查该共享文件夹的写入权限是否已开放给终端账号，截图 ai-init-folders 日志发给管理员。" % UID)
        else:
            item["msgs"].append("终端账号写入测试通过")
    except Exception as e:
        item["ok"] = False
        item["msgs"].append(f"初始化失败：{type(e).__name__}: {e}")
    results.append(item)
    print(("[OK]   " if item["ok"] else "[FAIL] ") + name + "：" + "；".join(item["msgs"]), flush=True)

os.makedirs(os.path.dirname(REPORT), exist_ok=True)
with open(REPORT, "w", encoding="utf-8") as f:
    json.dump({"time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "results": results}, f, ensure_ascii=False, indent=2)
print("部门文件夹初始化完成", flush=True)
