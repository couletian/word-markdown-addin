@echo off
chcp 65001 > nul
title 构建 Markdown工具箱.exe

echo.
echo ========================================
echo   构建单文件 exe
echo ========================================
echo.

cd /d "%~dp0"

REM 1) 检查 Python
where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.9+
    echo         https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM 2) 安装依赖
echo [1/4] 检查依赖...
python -c "import PyInstaller, cryptography" >nul 2>nul
if errorlevel 1 (
    echo       正在安装 PyInstaller 与 cryptography...
    python -m pip install --quiet --upgrade pyinstaller cryptography
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请检查网络
        pause
        exit /b 1
    )
)
echo       依赖就绪

REM 3) 生成图标
echo [2/4] 生成图标...
if not exist "assets\icon.ico" (
    python -c "import PIL" >nul 2>nul
    if errorlevel 1 (
        python -m pip install --quiet pillow
    )
    python assets\make-ico.py
) else (
    echo       assets\icon.ico 已存在
)

REM 4) 运行测试
echo [3/4] 运行测试...
node test\test-markdown.js >nul 2>nul
if errorlevel 1 (
    echo       [警告] 测试未通过或未安装 Node，继续构建
) else (
    echo       测试通过
)

REM 5) 打包
echo [4/4] 打包中（首次约需 1-2 分钟）...
python -m PyInstaller --clean --noconfirm --distpath dist --workpath build build_exe.spec >nul 2>nul
if errorlevel 1 (
    echo [错误] 打包失败
    echo        尝试手动运行：pyinstaller --clean --noconfirm build_exe.spec
    pause
    exit /b 1
)

echo.
echo ========================================
echo   构建完成
echo ========================================
echo.
if exist "dist\Markdown工具箱.exe" (
    for %%A in ("dist\Markdown工具箱.exe") do echo   输出：%%~fA
    for %%A in ("dist\Markdown工具箱.exe") do echo   大小：%%~zA 字节
) else (
    echo   [错误] 未找到输出文件
)
echo.
echo   双击 dist\Markdown工具箱.exe 即可运行
echo.
pause
