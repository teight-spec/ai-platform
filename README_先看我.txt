AI 助手平台 · 部署包
==================================================

完整步骤见部署手册（docx，不在代码仓库里），这里是最短路径：

1. DSM 建 4 个共享文件夹：AI资材、AI财务、AI营销、AI总裁办
   新建一个非管理员的服务账号给部门终端用，只给这 4 个文件夹「读取+写入」、docker 共享「只读」；
   在 NAS 上 id <账号名> 查到 uid/gid，填进 .env 的 TERM_UID / TERM_GID（不要给 Everyone 开权限）
2. 把 .env.example 复制为 .env，用记事本打开，只改前 3 行：ZHIPU_API_KEY、GATE_ADMIN_PASSWORD、GATE_VIEWER_PASSWORD（保存为 UTF-8）
3. 资源管理器地址栏输入 \\<NAS-IP>\docker ，把整个 ai-platform 文件夹复制进去
4. Container Manager → 项目 → 新增 → 路径选 /docker/ai-platform → 使用现有 docker-compose.yml
5. 打开 http://<NAS-IP>:8031 自检页，确认全绿
6. 打开 http://<NAS-IP>:8030 注册第一个账号（= 管理员）
7. 打开 http://<NAS-IP>:8031/admin → 测试智谱密钥 → 一键初始化 → 批量导入账号
8. 按手册第 4 章 P0 验证清单逐项测试

补充（2026-09-22）：部门终端改用办公版镜像 ai-term-office:1.1（Excel/Word/PPT/PDF 开源技能），
      获取方法见《本机测试说明.txt》第八节。
补充（2026-09-24）：按「数据处理 / 出图 / 周报 PPT」三类任务调整——公司工具包 company\kit、公司中文技能、
      斜杠快捷指令、单对话 ¥5 上限，详见《本机测试说明.txt》第九节。
补充（2026-09-24 v1.8）：部门说明与部门配方、月度价值页、配置检查、回归测试（本机测试工具菜单 C）、
      《运维卡.txt》（快照设置与恢复、故障处理），详见《本机测试说明.txt》第十节。
注意：.env 含密钥，不要外发。
