@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title AI 平台 - 本机测试工具

rem ================= 可改参数 =================
set "NAS_DIR=\\<NAS-IP>\docker\ai-platform"
rem ============================================

set "DC=docker compose -p ai-platform-local --env-file .env.local -f docker-compose.yml -f docker-compose.local.yml"
set "IMG1="
set "IMG2="
set "IMG3=ai-term-office:1.1"
for /f "usebackq tokens=1,* delims==" %%a in (".env.local") do (
  if "%%a"=="IMAGE_OPENWEBUI" set "IMG1=%%b"
  if "%%a"=="IMAGE_TERMINAL" set "IMG2=%%b"
  if "%%a"=="WEBUI_PORT" set "P_WEB=%%b"
  if "%%a"=="GATE_PORT" set "P_GATE=%%b"
)
if not defined P_WEB set "P_WEB=8030"
if not defined P_GATE set "P_GATE=8031"

:menu
cls
echo ==============================================================
echo    AI 平台 - 本机测试工具
echo ==============================================================
echo.
echo    1. 检查环境（Docker / 镜像 / 网络 / 端口）
echo    2. 启动本机平台（第一次或改了 docker-compose.yml 后）
echo    3. 改了代码后：重启费用闸门 + 部门终端
echo    4. 查看费用闸门日志
echo    5. 查看容器状态
echo    6. 停止本机平台（测试数据保留）
echo    7. 导入离线镜像（把 .tar 放进「镜像」文件夹）
echo    8. 同步到 NAS（先预览，确认后才复制）
echo    9. 打开网页（平台 localhost:%P_WEB% / 自检 localhost:%P_GATE%）
echo    A. 构建办公版终端镜像（改了 term-image 后，或第一次）
echo    B. 导出办公版终端镜像给 NAS（生成 .tar）
echo    C. 运行测试（公司工具包回归 + 各部门配方回归，改了 company 后先跑）
echo    0. 退出
echo.
set "c="
set /p "c=请输入数字后回车："
if "%c%"=="1" goto check
if "%c%"=="2" goto start
if "%c%"=="3" goto restart
if "%c%"=="4" goto logs
if "%c%"=="5" goto status
if "%c%"=="6" goto stop
if "%c%"=="7" goto import
if "%c%"=="8" goto sync
if "%c%"=="9" goto open
if /i "%c%"=="A" goto buildimg
if /i "%c%"=="B" goto saveimg
if /i "%c%"=="C" goto tests
if "%c%"=="0" exit /b 0
goto menu

rem ------------------------------------------------------------ 1 检查环境
:check
echo.
echo [1/5] Docker 引擎
docker version --format "      引擎版本 {{.Server.Version}}" 2>nul
if errorlevel 1 (
  echo       [未就绪] 没检测到 Docker。请先安装并打开 Rancher Desktop，
  echo                容器引擎选 dockerd（moby），等左下角状态变成绿色再试。
  goto done
)
docker compose version
echo.
echo [2/5] 本机已有的镜像
docker image inspect "%IMG1%" >nul 2>&1 && (echo       [有]   %IMG1%) || (echo       [没有] %IMG1%)
docker image inspect "%IMG2%" >nul 2>&1 && (echo       [有]   %IMG2%) || (echo       [没有] %IMG2%)
docker image inspect "%IMG3%" >nul 2>&1 && (echo       [有]   %IMG3%  （办公版终端）) || (echo       [没有] %IMG3%  （办公版终端，选 A 构建）)
echo.
echo [3/5] 网络（000 = 不通；200/401/404 等任意三位数 = 通）
call :probe ghcr.io/v2/ "ghcr.io（官方镜像源）"
call :probe ghcr.nju.edu.cn/v2/ "ghcr.nju.edu.cn（南大镜像站）"
call :probe open.bigmodel.cn/api/paas/v4/models "open.bigmodel.cn（智谱 API）"
echo.
echo [4/5] 端口占用
netstat -ano | findstr /r /c:":%P_WEB% .*LISTENING" >nul && (echo       %P_WEB% 平台端口 已被占用（如果是本平台在运行，属正常）) || (echo       %P_WEB% 平台端口 空闲)
netstat -ano | findstr /r /c:":%P_GATE% .*LISTENING" >nul && (echo       %P_GATE% 自检端口 已被占用（如果是本平台在运行，属正常）) || (echo       %P_GATE% 自检端口 空闲)
echo       （端口取自 .env.local；如被别的程序占用，改 .env.local 的 WEBUI_PORT 和 WEBUI_URL）
echo.
echo [5/5] 本机配置文件
for %%f in (.env.local docker-compose.yml docker-compose.local.yml) do (
  if exist "%%f" (echo       [有]   %%f) else (echo       [缺少] %%f)
)
echo.
echo 说明：两个镜像都「没有」且 ghcr.io 不通时 —— 按说明书第 3 节换镜像站或导入离线镜像。
goto done

:probe
set "code=000"
for /f %%h in ('curl -s -o nul -w "%%{http_code}" --max-time 10 https://%~1 2^>nul') do set "code=%%h"
echo       %code%   %~2
exit /b 0

rem ------------------------------------------------------------ C 测试
:tests
echo.
echo [1/2] 公司工具包回归测试（在资材终端里跑，只用 /tmp 临时文件）
%DC% exec -T term-zc python3 /opt/company/tests/test_kit.py
if errorlevel 1 (
  echo.
  echo [有问题] 上面有失败项，先别同步 NAS。平台没启动时先选 2。
  goto done
)
echo.
echo [2/2] 各部门配方回归（用配方自带的样例重跑，和验收值比对；没有配方的部门会显示「暂无」）
for %%t in (term-zc term-cw term-yx term-zb) do (
  echo ---- %%t
  %DC% exec -T %%t python3 /opt/company/kit/dept.py recipe-check
)
echo.
echo [完成] 全部通过再选 8 同步 NAS。
goto done

rem ------------------------------------------------------------ 2 启动
:start
for %%d in (AI资材 AI财务 AI营销 AI总裁办) do if not exist "nas-sim\%%d" mkdir "nas-sim\%%d"
if not exist "data-local" mkdir "data-local"
echo.
echo 正在启动……第一次需要下载约 6GB 镜像，可能要 30-60 分钟，窗口不要关。
echo.
%DC% up -d
if errorlevel 1 (
  echo.
  echo [失败] 请把上面的红字截图发给平台管理员；镜像下载失败见说明书第 3 节。
  goto done
)
echo.
echo [完成] 约 1 分钟后打开：
echo        自检页  http://localhost:%P_GATE%     （顶部橙色条 = 模拟模式）
echo        平台    http://localhost:%P_WEB%     （本机第一次打开要注册管理员）
echo        管理页  http://localhost:%P_GATE%/admin  密码见 .env.local
goto done

rem ------------------------------------------------------------ 3 重启
:restart
echo.
echo 重新跑部门文件夹初始化 ……
%DC% start init-folders
echo 重启费用闸门和四个部门终端 ……
%DC% restart cost-gate term-zc term-cw term-yx term-zb
if errorlevel 1 (echo [失败] 平台可能还没启动，先选 2。) else (echo [完成] 代码已生效。Open WebUI 没重启，页面刷新即可。)
goto done

rem ------------------------------------------------------------ 4 日志
:logs
echo.
echo 显示最近 100 行后持续刷新。看完按 Ctrl+C，再按 N 回到菜单。
echo.
%DC% logs -f --tail 100 cost-gate
goto done

rem ------------------------------------------------------------ 5 状态
:status
echo.
%DC% ps -a
echo.
echo 说明：ai-init-folders 显示 Exited (0) 是正常的（跑完一次就退出）。
goto done

rem ------------------------------------------------------------ 6 停止
:stop
echo.
%DC% down
echo.
echo [完成] 已停止。测试数据还在 data-local 和 nas-sim 文件夹里；
echo        想从零重测，停止后手动删掉这两个文件夹再选 2。
goto done

rem ------------------------------------------------------------ 7 导入镜像
:import
if not exist "镜像" mkdir "镜像"
if not exist "镜像\*.tar" (
  echo.
  echo 「镜像」文件夹里没有 .tar 文件。获取方法见说明书第 3 节。
  start "" "镜像"
  goto done
)
for %%f in ("镜像\*.tar") do (
  echo 导入 %%~nxf ……
  docker load -i "%%f"
)
goto done

rem ------------------------------------------------------------ 8 同步到 NAS
:sync
echo.
if not exist "%NAS_DIR%\docker-compose.yml" (
  echo [失败] 访问不到 %NAS_DIR%
  echo        先在资源管理器地址栏打开 \\<NAS-IP>\docker 登录一次，再回来重试。
  goto done
)
set "RC=robocopy "%~dp0." "%NAS_DIR%" /E /XO /R:1 /W:2 /NJH /NDL /NP /FP /XD "%~dp0data" "%~dp0data-local" "%~dp0nas-sim" "%~dp0镜像" __pycache__ .git /XF .env .env.local docker-compose.local.yml 本机测试工具.bat 本机测试说明.txt *.pyc *.zip *.docx"
echo 预览：下面这些文件会被复制到 NAS（只增不删；NAS 上更新的文件会跳过）
echo 不会碰：NAS 的 data 文件夹、.env（密钥）、部门共享文件夹
echo --------------------------------------------------------------
%RC% /L
echo --------------------------------------------------------------
set "ok="
set /p "ok=确认复制请输入 Y 回车（其他键取消）："
if /i not "%ok%"=="Y" (echo 已取消。& goto done)
%RC%
if errorlevel 8 (echo [失败] 复制出错，请截图。& goto done)
echo.
echo [完成] 已同步。到 NAS 的 Container Manager 里让改动生效：
echo   - 只改了 cost-gate 里的 .py      → 容器 → ai-cost-gate → 重新启动
echo   - 改了 company 文件夹           → 容器 → ai-init-folders 启动一次，再重启 ai-term-* 四个
echo   - 改了 docker-compose.yml       → 项目 → ai-platform → 停止 → 构建
echo   - NAS 的 .env 不会被同步；要改就直接编辑 %NAS_DIR%\.env
echo   - 确认 NAS 自检页顶部【没有】橙色「模拟模式」条
goto done

rem ------------------------------------------------------------ A 构建办公版终端镜像
:buildimg
echo.
echo 构建办公版终端镜像 %IMG3% ……
echo   基于 %IMG2%，补装 xlrd / pdfplumber / reportlab / pptxgenjs / 中文字体
echo   第一次约 5-15 分钟（走阿里云、npmmirror 国内源），窗口不要关。
echo.
%DC% build term-zc
if errorlevel 1 (
  echo.
  echo [失败] 请把上面的红字截图发给平台管理员。
  goto done
)
echo.
echo 让四个部门终端换用新镜像 ……
%DC% up -d --no-deps term-zc term-cw term-yx term-zb
echo.
echo [完成] 自检：
docker exec ai-term-zc sh /opt/company/selfcheck.sh
goto done

rem ------------------------------------------------------------ B 导出镜像
:saveimg
if not exist "镜像" mkdir "镜像"
docker image inspect "%IMG3%" >nul 2>&1 || (echo. & echo 本机还没有 %IMG3%，先选 A 构建。& goto done)
echo.
echo 导出 %IMG3% 到「镜像\ai-term-office_1.1.tar」（约 3GB，需要几分钟）……
docker save -o "镜像\ai-term-office_1.1.tar" %IMG3%
if errorlevel 1 (echo [失败] 导出出错，请截图。& goto done)
echo.
echo [完成] NAS 上导入：Container Manager → 映像 → 新增 → 从文件添加 → 选这个 .tar
echo        导入后：项目 → ai-platform → 停止 → 构建（会直接用导入的镜像，不再下载）
start "" "镜像"
goto done

rem ------------------------------------------------------------ 9 打开网页
:open
start "" http://localhost:%P_GATE%
start "" http://localhost:%P_WEB%
goto menu

:done
echo.
pause
goto menu
