@echo off
echo 正在安装PSP ISO文件排序工具...
echo.

REM 检查conda是否已安装
conda --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到conda，请先安装Anaconda或Miniconda
    echo 下载地址：https://www.anaconda.com/products/distribution
    pause
    exit /b 1
)

echo Conda已找到
conda --version

echo.
echo 正在创建conda环境...
conda create -n psp-iso-sorter python=3.9 -y

echo.
echo 正在激活环境并安装依赖包...
call conda activate psp-iso-sorter
pip install pywin32

if errorlevel 1 (
    echo.
    echo 依赖安装可能失败，但程序仍可运行
    echo 注意：pywin32包用于更好的Windows文件时间控制，如果安装失败程序将使用备用方法
)

echo.
echo 安装完成！
echo.
echo 使用方法：
echo   双击运行 run.bat 或者
echo   在Anaconda Prompt中运行: 
echo     conda activate psp-iso-sorter
echo     python psp_iso_sorter.py
echo.
pause
