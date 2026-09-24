# AI 助手平台（部门版）

部署在内网 NAS 上的部门 AI 助手：员工网页登录，按部门隔离，用 AI 做数据处理、出图表、周报汇报（交付 PDF）。

- **Open WebUI**：登录、对话界面、部门分组与权限、斜杠快捷指令
- **Open Terminal**（每个部门一个容器）：AI 在这里执行 Python / LibreOffice，只挂本部门文件夹、断外网、非 root 运行
- **费用闸门 cost-gate**（本仓库自研）：所有模型调用的唯一出口，按部门计费与限额、单对话上限、自检页与管理页、一键初始化
- **company/**：公司工具包（数据摘要、公司配色出图、公司格式 Excel、可选的周报出数工具）、中文技能、开源 Excel/Word/PPT/PDF 技能（来源与许可见 `company/skills/来源与许可.md`）

快速开始见 `README_先看我.txt`；本机（Rancher Desktop）模拟测试见 `本机测试说明.txt`。
配置模板：`.env.example`（复制为 `.env` 后填写；`.env` 不进仓库）。
