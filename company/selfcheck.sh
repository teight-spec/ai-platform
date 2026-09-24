#!/bin/sh
# 部门终端自检：docker exec ai-term-zc sh /opt/company/selfcheck.sh
ok=0; bad=0
chk() { if sh -c "$2" >/dev/null 2>&1; then echo "  [OK]   $1"; ok=$((ok+1)); else echo "  [缺少] $1"; bad=$((bad+1)); fi; }
echo "部门终端自检："
chk "Python 办公组件（xlrd/pdfplumber/reportlab/markitdown/pdf2image）" "python3 -c 'import xlrd,pdfplumber,reportlab,markitdown,pdf2image'"
chk "Python 基础组件（pandas/openpyxl/docx/pptx/pypdf/matplotlib）" "python3 -c 'import pandas,openpyxl,docx,pptx,pypdf,matplotlib'"
chk "Node pptxgenjs" "node -e \"require('pptxgenjs')\""
chk "LibreOffice（soffice）" "command -v soffice"
chk "poppler（pdftoppm/pdftotext）" "command -v pdftoppm && command -v pdftotext"
chk "中文字体替换（微软雅黑→文泉驿）" "fc-match 'Microsoft YaHei' | grep -qi wenquanyi"
chk "开源技能目录 /opt/company/skills" "test -f /opt/company/skills/INDEX.md"
chk "转 PDF 脚本 office2pdf.py" "test -f /opt/company/office2pdf.py"
chk "部门说明与配方工具（dept.py/runio）" "python3 /opt/company/kit/dept.py --help && cd /tmp && python3 -c 'import sys; sys.path.insert(0,\"/opt/company/kit\"); import runio'"
chk "公司工具包（peek/charts/tables/weekly）" "python3 /opt/company/kit/weekly.py --help && cd /tmp && python3 -c 'import sys; sys.path.insert(0,\"/opt/company/kit\"); import charts, tables'"
chk "终端守护 term_guard（命令超时终止、清理临时文件）" "grep -aqs 'term_guar[d].py' /proc/[0-9]*/cmdline"
chk "本部门文件夹可读写（当前账号 $(id -un) uid=$(id -u)）" "ls . && test -w ../02_输出"
echo "结果：通过 $ok 项，缺少 $bad 项"
[ "$bad" -eq 0 ]
