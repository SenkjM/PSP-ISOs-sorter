@echo off
title PSP ISO文件排序工具
echo 启动PSP ISO文件排序工具...
echo.

REM 检查conda环境是否存在
conda env list | findstr "psp-iso-sorter" >nul
if errorlevel 1 (
    echo 错误：未找到psp-iso-sorter环境，请先运行install.bat
    pause
    exit /b 1
)

REM 激活conda环境并运行程序
call conda activate psp-iso-sorter
python psp_iso_sorter.py

if errorlevel 1 (
    echo.
    echo 程序运行出错，请检查：
    echo 1. 是否运行了install.bat安装依赖
    echo 2. conda环境是否正确配置
    echo.
    pause
)
