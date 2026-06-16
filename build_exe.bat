@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  กำลังสร้างไฟล์ .exe (ใช้เวลาสักครู่ 2-5 นาที)
echo ============================================================
py -3.12 -m pip install --upgrade pyinstaller
py -3.12 -m PyInstaller --noconfirm --onefile --windowed ^
  --name "TriangleSign" ^
  --icon "app.ico" ^
  --splash "splash.png" ^
  --add-data "templates;templates" ^
  --add-data "fonts;fonts" ^
  --add-data "sample_list.xlsx;." ^
  --add-data "app.ico;." ^
  --hidden-import win32timezone ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32print ^
  app.py
echo.
echo ============================================================
echo  เสร็จแล้ว! ไฟล์โปรแกรมอยู่ที่   dist\TriangleSign.exe
echo  (ดับเบิลคลิกเปิดได้เลย ไม่ต้องลง Python)
echo ============================================================
pause
