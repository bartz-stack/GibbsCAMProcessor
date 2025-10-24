@echo off
REM ============================================================================
REM Build GibbsCAM Processor Executable
REM ============================================================================
REM This script builds a standalone .exe file using PyInstaller
REM
REM Requirements:
REM   - Python with all dependencies installed
REM   - PyInstaller: pip install pyinstaller
REM   - Gibbscam.ico file (copy from network to GCP_clau folder)
REM
REM Output: dist\GibbsCAMProcessor.exe
REM ============================================================================
echo.
echo ============================================================================
echo GibbsCAM Processor - Build Script
echo ============================================================================
echo.
REM Check if icon exists locally
if not exist "GCP_clau\Gibbscam.ico" (
    echo WARNING: Gibbscam.ico not found in GCP_clau folder
    echo.
    echo Please copy the icon file to GCP_clau\Gibbscam.ico
    echo Source: \\Schuette-DC1\production\GibbsCAM\NPD Setup Sheets\Script Files\Gibbscam.ico
    echo.
    pause
    exit /b 1
)
REM Clean old build files
echo [1/4] Cleaning old build files...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "GibbsCAMProcessor.spec" del GibbsCAMProcessor.spec
echo   Done.
echo.
REM Install/upgrade PyInstaller
echo [2/4] Checking PyInstaller...
pip install --upgrade pyinstaller >nul 2>&1
echo   Done.
echo.
REM Build the executable
echo [3/4] Building executable (this may take a minute)...
pyinstaller ^
    --name=GibbsCAMProcessor ^
    --onefile ^
    --windowed ^
    --icon=GCP_clau\Gibbscam.ico ^
    --add-data "GCP_clau\config.ini;GCP_clau" ^
    --add-data "GCP_clau\Gibbscam.ico;GCP_clau" ^
    --hidden-import=win32com.client ^
    --hidden-import=win32gui ^
    --hidden-import=win32process ^
    --hidden-import=psutil ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=winotify ^
    --hidden-import=PIL._tkinter_finder ^
    --hidden-import=PIL.Image ^
    --hidden-import=PIL.ImageTk ^
    --hidden-import=PIL.ImageGrab ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.messagebox ^
    GCP_clau\__main__.py
if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)
echo   Done.
echo.
REM Check if executable was created
echo [4/4] Verifying executable...
if exist "dist\GibbsCAMProcessor.exe" (
    echo   SUCCESS! Executable created.
    echo.
    echo ============================================================================
    echo Build Complete!
    echo ============================================================================
    echo.
    echo Executable location: dist\GibbsCAMProcessor.exe
    echo.
    echo You can now:
    echo   1. Test the exe: dist\GibbsCAMProcessor.exe
    echo   2. Copy to your desired location
    echo   3. Create a desktop shortcut
    echo.
    echo The config.ini and icon are embedded in the executable.
    echo.
) else (
    echo   ERROR: Executable was not created!
    echo.
    pause
    exit /b 1
)
pause