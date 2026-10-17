@echo off
chcp 65001 > nul
title RED TEAM: The Adversary Terminal CLI - DT005
echo =========================================================================
echo   Đang khởi động RED TEAM: The Adversary Terminal Console...
echo   Mục tiêu tấn công: Blue Team (Thư mục giám sát)
echo =========================================================================
python src/adversary/red_team_cli.py
pause
