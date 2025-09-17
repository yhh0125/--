@echo off
chcp 65001 >nul
title 房租计算程序 - 一键启动
color 0A

echo ========================================
echo           房租计算程序
echo ========================================
echo.
echo 正在启动程序...
echo.

cd /d "C:\Users\lenovo\Desktop\房租计算"

echo 启动Flask应用...
start /min py app.py

echo 等待程序启动...
timeout /t 3 /nobreak >nul

echo 正在打开浏览器...
start http://127.0.0.1:5000

echo.
echo 程序已启动！
echo 浏览器已自动打开
echo.
echo 按任意键关闭此窗口...
pause >nul

