@echo off
setlocal EnableDelayedExpansion
title Invoice Splitter - Setup
color 1F
echo.
echo  ============================================
echo   Invoice Splitter - Setup
echo  ============================================
echo.
echo   [1]  Fresh Install  (new computer)
echo   [2]  Update
echo   [3]  Exit
echo.
choice /C 123 /N /M "  Choose 1, 2 or 3: "
set CHOICE=%errorlevel%

if "%CHOICE%"=="1" goto INSTALL
if "%CHOICE%"=="2" goto UPDATE
if "%CHOICE%"=="3" exit /b 0
echo   Invalid choice.
pause & exit /b 1


:: ===========================================================
:INSTALL
:: ===========================================================
echo.
echo  ============================================
echo   התקנה חדשה
echo  ============================================
echo.

:: Check admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [!] יש להריץ כמנהל מערכת.
    echo     לחץ ימני על setup.bat ובחר "Run as administrator"
    pause & exit /b 1
)

:: Python
echo [1/4] בודק Python...
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Python מותקן.
    goto python_done
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto python_done
)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto python_done
)
echo מוריד Python 3.11...
curl -L --output "%TEMP%\python_installer.exe" "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
if not exist "%TEMP%\python_installer.exe" ( echo שגיאה בהורדת Python. & pause & exit /b 1 )
echo מתקין Python...
"%TEMP%\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_tcltk=1
timeout /t 10 /nobreak >nul
set "PYTHON="
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if "!PYTHON!"=="" ( echo שגיאה: לא נמצא Python לאחר ההתקנה. & pause & exit /b 1 )

:python_done
if "!PYTHON!"=="" set "PYTHON=python"

:: Tesseract
echo.
echo [2/4] בודק Tesseract OCR...
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo [OK] Tesseract מותקן.
    goto tesseract_done
)
echo מוריד Tesseract OCR...
curl -L --output "%TEMP%\tesseract_setup.exe" "https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe"
if not exist "%TEMP%\tesseract_setup.exe" ( echo שגיאה בהורדת Tesseract. & pause & exit /b 1 )
echo מתקין Tesseract...
"%TEMP%\tesseract_setup.exe" /S /D=C:\Program Files\Tesseract-OCR
timeout /t 8 /nobreak >nul
if not exist "C:\Program Files\Tesseract-OCR\tesseract.exe" ( echo שגיאה בהתקנת Tesseract. & pause & exit /b 1 )
echo [OK] Tesseract מותקן.

:tesseract_done

:: pip + packages
echo.
echo [3/4] מעדכן pip...
"!PYTHON!" -m pip install --upgrade pip --quiet

echo.
echo [4/4] מתקין ספריות Python...
"!PYTHON!" -m pip install pymupdf pytesseract pillow numpy --quiet
if %errorlevel% neq 0 ( echo שגיאה בהתקנת ספריות. & pause & exit /b 1 )
echo [OK] ספריות מותקנות.

goto VERIFY


:: ===========================================================
:UPDATE
:: ===========================================================
echo.
echo  ============================================
echo   UPDATE - starting...
echo  ============================================
echo.
pause

:: Find Python
set "PYTHON="
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if exist "C:\Python311\python.exe" set "PYTHON=C:\Python311\python.exe"
if exist "C:\Python312\python.exe" set "PYTHON=C:\Python312\python.exe"
if exist "C:\Python313\python.exe" set "PYTHON=C:\Python313\python.exe"
python --version >nul 2>&1
if %errorlevel% equ 0 if "!PYTHON!"=="" set "PYTHON=python"
echo Python path: [!PYTHON!]
if "!PYTHON!"=="" ( echo שגיאה: Python לא נמצא. הרץ התקנה חדשה תחילה. & pause & exit /b 1 )
echo [OK] Python נמצא.

:: Download latest invoice_splitter.py from GitHub
echo.
echo [1/2] מוריד גרסה עדכנית של התוכנה...
curl -L --output "%~dp0invoice_splitter.py" "https://raw.githubusercontent.com/mici2854911-ship-it/invoice-splitter/master/invoice_splitter.py"
echo curl result: %errorlevel%
if %errorlevel% neq 0 ( echo שגיאה בהורדת העדכון. בדוק חיבור לאינטרנט. & pause & exit /b 1 )
echo [OK] התוכנה עודכנה.
pause

:: Update packages
echo.
echo [2/2] מעדכן ספריות Python...
"!PYTHON!" -m pip install --upgrade pip --quiet
"!PYTHON!" -m pip install --upgrade pymupdf pytesseract pillow numpy --quiet
if %errorlevel% neq 0 ( echo שגיאה בעדכון ספריות. & pause & exit /b 1 )
echo [OK] ספריות עודכנו.

goto VERIFY


:: ===========================================================
:VERIFY
:: ===========================================================
echo.
echo  ============================================
echo   בדיקה
echo  ============================================
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo [OK] Tesseract נמצא.
) else (
    echo [WARN] Tesseract לא נמצא - OCR לא יעבוד.
)
"!PYTHON!" -c "import fitz, pytesseract, PIL, numpy, tkinter; print('[OK] כל הספריות מוכנות.')"
if %errorlevel% neq 0 ( echo [FAIL] ספריות חסרות. & pause & exit /b 1 )

echo.
echo  ============================================
if "!CHOICE!"=="1" (
    echo   ההתקנה הושלמה בהצלחה!
) else (
    echo   העדכון הושלם בהצלחה!
)
echo  ============================================
echo.
echo   להפעלה: לחץ כפתור ימני על invoice_splitter.py
echo            ובחר "Open With Python"
echo.
pause
