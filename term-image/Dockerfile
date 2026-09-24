# ==========================================================================
#  部门终端镜像（办公增强版）= Open Terminal 0.13.0 + 办公文件处理补充组件
#
#  为什么要自建：部门终端断外网，运行时装不了包，必须提前做进镜像。
#  基础镜像已自带：LibreOffice、poppler(pdftoppm)、pandoc、Node 22、
#                   pandas、openpyxl、python-docx、python-pptx、pypdf、matplotlib
#  本镜像补充：
#    xlrd        读老格式 .xls（ERP 导出常见）
#    pdfplumber  提取 PDF 表格
#    reportlab   生成 PDF；stamp_internal.py 给 PDF 加标注也要用它
#    pdf2image   PDF 转图片检查排版（doc 技能的 render_docx.py 用）
#    markitdown  PPT/Word/Excel 快速转文字（pptx-generator 技能用）
#    pptxgenjs   Node 生成 PPT（pptx-generator 技能用）
#    文泉驿字体 + 字体替换表：微软雅黑/宋体/黑体/等线 → 文泉驿正黑，转 PDF 不缺字
#
#  构建（本机测试工具.bat 菜单 A 会自动执行）：
#    docker build -t ai-term-office:1.1 term-image
#  国内镜像源默认开着；在有外网直连的地方构建时，可以把三个 MIRROR 参数传空值。
# ==========================================================================
ARG BASE_IMAGE=ghcr.io/open-webui/open-terminal:0.13.0
FROM ${BASE_IMAGE}

ARG APT_MIRROR=mirrors.aliyun.com
ARG PIP_INDEX=https://mirrors.aliyun.com/pypi/simple/
# 阿里云缺版本时自动到清华源找（阿里云 PyPI 同步有滞后，2026-09-22 就缺 markitdown 0.1.8）
ARG PIP_EXTRA_INDEX=https://pypi.tuna.tsinghua.edu.cn/simple/
ARG NPM_REGISTRY=https://registry.npmmirror.com

USER root

# 1) 中文字体（系统包，约 10MB）
RUN set -eux; \
    if [ -n "${APT_MIRROR}" ]; then \
      sed -i "s|deb.debian.org|${APT_MIRROR}|g" /etc/apt/sources.list.d/*.sources 2>/dev/null || true; \
      sed -i "s|deb.debian.org|${APT_MIRROR}|g" /etc/apt/sources.list 2>/dev/null || true; \
    fi; \
    apt-get update; \
    apt-get install -y --no-install-recommends fonts-wqy-zenhei fontconfig; \
    rm -rf /var/lib/apt/lists/*

COPY fonts-cjk-alias.conf /etc/fonts/local.conf

# 2) Python 补充包（固定版本，保证 NAS 和本机一致）
RUN set -eux; \
    if [ -n "${PIP_INDEX}" ]; then export PIP_INDEX_URL="${PIP_INDEX}"; fi; \
    if [ -n "${PIP_EXTRA_INDEX}" ]; then export PIP_EXTRA_INDEX_URL="${PIP_EXTRA_INDEX}"; fi; \
    pip install --no-cache-dir \
      xlrd==2.0.2 \
      pdfplumber==0.11.10 \
      reportlab==5.0.1 \
      pdf2image==1.17.0 \
      "markitdown[pptx,docx,xlsx,xls,pdf]==0.1.7"

# 3) Node 补充包
RUN set -eux; \
    if [ -n "${NPM_REGISTRY}" ]; then npm config set registry "${NPM_REGISTRY}"; fi; \
    npm install -g pptxgenjs@4.0.1; \
    npm cache clean --force; \
    fc-cache -f

# 让 node 脚本能直接 require('pptxgenjs')
ENV NODE_PATH=/usr/lib/node_modules

# 容器内运行账号 user 的 UID/GID 改成群晖终端服务账号 的 UID/GID
# 这样 NAS 共享文件夹只需授权给 终端服务账号（不用开 Everyone），容器照样以普通账号运行、不用 root
# 在 NAS 上用 `id <服务账号>` 查到数字后填在这里（群晖普通用户的主组一般是 users = 100）
ARG TERM_UID=1000
ARG TERM_GID=100
RUN set -eux; \
    if [ "$(id -u user)" != "${TERM_UID}" ] || [ "$(id -g user)" != "${TERM_GID}" ]; then \
      getent group "${TERM_GID}" >/dev/null || groupadd -g "${TERM_GID}" aiterm; \
      usermod -u "${TERM_UID}" -g "${TERM_GID}" user; \
      chown -R "${TERM_UID}:${TERM_GID}" /home/user; \
    fi; \
    id user

# 去掉原版给 user 的免密 sudo：AI 生成的代码不能在容器里提权成 root
# （原版入口脚本里的 sudo 只在装包/挂 docker.sock 时用到，本平台都不用；失败会被忽略）
RUN sed -i '/^user ALL=(ALL) NOPASSWD:ALL/d' /etc/sudoers && ! grep -q '^user ALL' /etc/sudoers

# 4) 自检：任何一项缺失，构建直接失败
RUN set -eux; \
    python3 -c "import xlrd, pdfplumber, reportlab, pdf2image, markitdown, openpyxl, docx, pptx, pypdf, pandas, matplotlib"; \
    node -e "require('pptxgenjs')"; \
    command -v soffice; command -v pdftoppm; command -v pdftotext; \
    fc-match "Microsoft YaHei" | grep -i "WenQuanYi"

USER user
WORKDIR /home/user
