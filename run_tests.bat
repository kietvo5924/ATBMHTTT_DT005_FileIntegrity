@echo off
chcp 65001 >nul
echo =========================================================================
echo       CHAY BO KIEM THU TU DONG TOAN DIEN DT005 (17 TESTS)
echo =========================================================================
python -m unittest discover -s src/tests -v
if %ERRORLEVEL% equ 0 (
    echo.
    echo [+] TAT CA 17/17 BAI KIEM THU DA PASS THANH CONG!
) else (
    echo.
    echo [-] CO BAI KIEM THU THAT BAI!
)
pause
