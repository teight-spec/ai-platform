# 技能总目录（处理 Excel / Word / PPT / PDF 前先读本文件）


## 一、按任务选技能

**先看公司技能（中文，已按本公司环境写好），再看开源技能。**

| 任务 | 读哪个文件 |
|---|---|
| 整理/清洗数据、汇总统计、出图、出 Excel 汇总表 | `/opt/company/skills/公司-数据处理与出图/SKILL.md` |
| 周报（先了解部门现在的做法，再照做 → PDF） | `/opt/company/skills/公司-周报PPT/SKILL.md` |
| 记住口径、存成配方、按上次的做法再做、用部门配方 | `/opt/company/skills/公司-部门说明与配方/SKILL.md` |
| Excel 的特殊操作：改公式、改格式、在原文件上插行插列、公式检查 | `/opt/company/skills/minimax-xlsx/SKILL.md` |
| 专题汇报 PPT、周报的 PPT 制作 | `/opt/company/skills/pptx-generator/SKILL.md` |
| 新建或修改 Word（.docx） | `/opt/company/skills/doc/SKILL.md` |
| 读取 PDF、提取 PDF 里的表格、合并/拆分 PDF、生成 PDF | `/opt/company/skills/pdf/SKILL.md` |

部门自己的技能和配方在 `~/<部门>/04_技能/` 里（`dept.py start` 会列出来），**部门技能和本目录冲突时，以部门技能为准**。
一次只读需要的那一个 SKILL.md；SKILL.md 里提到的 references/*.md 只在确实需要时再读。

### 公司工具包 `/opt/company/kit/`（直接调用，不要复制改写）

| 工具 | 用途 |
|---|---|
| `python3 /opt/company/kit/dept.py start` | **任务开始先跑**：部门说明（已确认的口径）+ 部门配方清单 |
| `from runio import io_args, save_summary` | script.py 可重跑写法：接收 --input/--out，写 结果摘要.json |
| `python3 /opt/company/kit/peek.py 文件` | 数据文件摘要（行数、列类型、空值、样例、可能的问题）。**任何 Excel/CSV 先跑它** |
| `from charts import line, bar, bar_compare, combo` | 公司配色出图（先 `sys.path.insert(0, "/opt/company/kit")`） |
| `from tables import save_excel` | 公司格式 Excel（列宽撑开、表头、冻结、内部标注） |
| `python3 /opt/company/kit/weekly.py` | 可选：数字表格型周报的固定出数工具，部门同意才用，见「公司-周报PPT」末尾 |

## 二、公司环境规定（优先于技能原文）

技能是开源的英文原版，写给有外网的环境。在本公司终端里，下面这些规定**覆盖**技能原文：

1. **不安装任何东西。** 终端没有外网。技能里的 `pip install`、`npm install`、`uv`、`brew`、`apt-get` 一律跳过。需要的都已装好：
   - Python：pandas、openpyxl、xlrd、python-docx、python-pptx、pypdf、pdfplumber、reportlab、pdf2image、markitdown、matplotlib
   - Node：pptxgenjs（直接 `require('pptxgenjs')`）
   - 命令：soffice（LibreOffice）、pdftoppm、pdftotext（poppler）、pandoc
   - 缺了就如实告诉用户「终端缺少 xxx」，不要想办法安装。
2. **路径。** 技能里的 `scripts/xxx.py` 指 `/opt/company/skills/<技能名>/scripts/xxx.py`。这个目录只读，临时文件写到 `/tmp/`。
3. **输出目录。** 用系统提示词规定的部门输出目录，不用技能里写的 `output/`、`slides/output/` 等目录。
4. **配色。** 用公司配色，不用技能里的配色表：蓝 `2F6FB0`、绿 `3A9A5B`、橙 `D9822B`、灰 `8C8C8C`，深色文字 `262626`，浅底色 `F5F7FA`。
   pptx-generator 的 theme 对象固定写成：
   `{primary:"2F6FB0", secondary:"3A9A5B", accent:"D9822B", light:"8C8C8C", bg:"FFFFFF"}`
5. **字体。** PPT、Word 的中文字体写「Microsoft YaHei」。终端里没有这个字体，转 PDF 时会自动换成文泉驿正黑，这是正常的。matplotlib 已配好中文字体。
6. **pptx-generator 的特别说明：**
   - 第 1 步「Search」改为：直接问用户汇报对象、页数、重点。
   - 不要用 react-icons / sharp 做图标（没有安装），用简单形状代替。
   - 第 5 步提到的 subagents 不存在，自己按顺序逐页写。
7. **PPT 只交 PDF。** 生成 .pptx 后运行
   `python3 /opt/company/office2pdf.py <文件.pptx>`
   得到同名 PDF。给用户的只有 PDF；.pptx 会被自动移到同目录的「底稿」子文件夹。Word、Excel 需要 PDF 版时也用这个命令（原文件不动）。
8. **检查成品。** 你可能看不到图片。用 `pdftotext -layout 文件.pdf -` 抽出文字，检查有没有缺字、乱码、文字溢出、页数不对。
   Excel 有公式时：先 `python3 /opt/company/skills/minimax-xlsx/scripts/libreoffice_recalc.py 文件.xlsx /tmp/重算.xlsx`，再对 `/tmp/重算.xlsx` 跑同目录的 `formula_check.py`，报错（#DIV/0! 等）要改掉再交付。
9. **最后一步。** 交付前运行 `python3 /opt/company/stamp_internal.py <输出目录> -r`，加「内部文件，禁止外传」标注。

## 三、常见情况

0. 临时文件用 `mktemp -d` 建自己的临时目录（同部门可能有别人同时在用终端），不要写固定文件名到 /tmp。
1. 读 Excel 先跑 `peek.py`，不要用 pandas 把整张表打印出来；大文件只打印汇总结果（不超过 30 行）。
2. 金额、数量列读进来是文本时，先转成数字，转不了的列出来给用户看。
3. 中文 CSV 读不出来，试 `encoding="gbk"`。
4. 老格式 .xls：`pd.read_excel(路径, sheet_name=None, engine="xlrd")`。需要改写时，另存为 .xlsx 再处理，原文件不动。
5. PDF 里的表格：用 pdfplumber 的 `page.extract_tables()`。
6. PDF 抽不出文字（扫描件、图片）：终端没有文字识别功能，如实告诉用户「这是扫描件，无法自动读取」，请用户提供电子版。
