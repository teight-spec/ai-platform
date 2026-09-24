#!/bin/bash
# ==========================================================================
#  部门终端启动包装
#  docker-compose.yml 里部门终端的 entrypoint 指向本脚本（镜像不用重建）：
#    1) 后台启动守护 term_guard.py：命令超过 TERM_CMD_MAX_MIN 分钟自动终止；每小时清理 /tmp 旧临时文件
#    2) exec 原版入口 /app/entrypoint.sh（Open Terminal 自带），参数原样传过去
#  守护出问题不影响终端本身；要临时关掉守护，在 compose 里设 TERM_GUARD=0。
# ==========================================================================
if [ "${TERM_GUARD:-1}" != "0" ] && [ -f /opt/company/term_guard.py ]; then
    python3 /opt/company/term_guard.py &
fi
exec /app/entrypoint.sh "$@"
