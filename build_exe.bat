@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem อ่านเลขเวอร์ชันจาก app.py (APP_VERSION = "x.y.z") มาตั้งชื่อไฟล์ให้อัตโนมัติ
set "VER=dev"
for /f tokens^=2^ delims^=^" %%v in ('findstr /b /c:"APP_VERSION" app.py') do set "VER=%%v"

echo ============================================================
echo  กำลังสร้างไฟล์ .exe เวอร์ชัน %VER% (ใช้เวลาสักครู่ 2-5 นาที)
echo ============================================================
py -3.12 -m pip install --upgrade pyinstaller pillow

rem สร้างรูปหน้าโหลด (splash) ให้ตรงเวอร์ชันล่าสุดก่อนเสมอ
py -3.12 make_splash.py

py -3.12 -m PyInstaller --noconfirm --onefile --windowed ^
  --name "TriangleSign-v%VER%" ^
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
echo  เสร็จแล้ว! ไฟล์โปรแกรมอยู่ที่   dist\TriangleSign-v%VER%.exe
echo  (ดับเบิลคลิกเปิดได้เลย ไม่ต้องลง Python)
echo ============================================================
pause
