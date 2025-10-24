@echo off
REM ============================================================================
REM Build GibbsCAM Processor - WITH ALL MODULES EXPLICITLY INCLUDED
REM ============================================================================
echo.
echo ============================================================================
echo GibbsCAM Processor - Build Script
echo ============================================================================
echo.

REM Clean old build files
echo [1/3] Cleaning old build files...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del *.spec
echo   Done.
echo.

REM Build with ALL modules as additional Python files
echo [2/3] Building executable with ALL modules...
echo.

pyinstaller ^
    --onefile ^
    --windowed ^
    --name=GibbsCAMProcessor ^
    --icon=Gibbscam.ico ^
    --add-data "config.ini;." ^
    --add-data "Gibbscam.ico;." ^
    --hidden-import=win32com.client ^
    --hidden-import=win32gui ^
    --hidden-import=win32process ^
    --hidden-import=psutil ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=windows_toasts ^
    --hidden-import=PIL ^
    --hidden-import=PIL.Image ^
    --hidden-import=PIL.ImageTk ^
    --hidden-import=PIL.ImageGrab ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.messagebox ^
    --paths=. ^
    processor.py ^
    config.py ^
    logging_setup.py ^
    notifications.py ^
    ncf_parser.py ^
    excel_mapper.py ^
    window_detector.py ^
    screenshot_gui.py ^
    screenshot_capture.py ^
    screenshot_colors.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)
echo   Done.
echo.

REM Verify
echo [3/3] Verifying executable...
if exist "dist\GibbsCAMProcessor.exe" (
    echo   SUCCESS!
    echo.
    echo Executable: dist\GibbsCAMProcessor.exe
    dir "dist\GibbsCAMProcessor.exe" | find "GibbsCAMProcessor.exe"
    echo.
) else (
    echo   ERROR: Executable not created!
    pause
    exit /b 1
)
pause