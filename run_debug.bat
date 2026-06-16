@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === โหมด debug: ถ้าโปรแกรมปิดเอง จะเห็นข้อความ error ที่นี่ ===
py -3.12 app.py
echo.
pause
