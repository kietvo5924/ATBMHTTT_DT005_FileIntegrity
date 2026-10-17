@echo off
chcp 65001 > nul
title RED TEAM: The Adversary Web Attack Console (Port 8081) - DT005
echo =========================================================================
echo   Đang khởi động RED TEAM: The Adversary Web Dashboard (Port 8081)...
echo   Mục tiêu tấn công: Blue Team (http://127.0.0.1:8080)
echo   Mở trình duyệt Red Team tại: http://127.0.0.1:8081
echo =========================================================================
start http://127.0.0.1:8081
python src/adversary/red_team_web.py --port 8081
pause
