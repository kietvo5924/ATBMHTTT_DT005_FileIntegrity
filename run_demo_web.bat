@echo off
chcp 65001 > nul
title Hệ Thống Hợp Nhất Web Demo & Benchmark Dashboard - DT005
echo =========================================================================
echo   Đang khởi động Hệ Thống Hợp Nhất Web Demo & Benchmark DT005...
echo   Mở trình duyệt tại: http://127.0.0.1:8080
echo   Bao gồm: Live Demo 4 Mô Hình + Bảng Điều Khiển Biểu Đồ Benchmark Động
echo =========================================================================
start http://127.0.0.1:8080
python src/demo/web_demo.py
pause
