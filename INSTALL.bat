@echo off
echo GibbsCAM Processor - Installation
echo ==================================
echo.
echo This will install GibbsCAM Processor to:
echo %USERPROFILE%\GibbsCAMProcessor
echo.
pause

mkdir "%USERPROFILE%\GibbsCAMProcessor"
copy GibbsCAMProcessor.exe "%USERPROFILE%\GibbsCAMProcessor\"
copy config.ini "%USERPROFILE%\GibbsCAMProcessor\"
copy Gibbscam.ico "%USERPROFILE%\GibbsCAMProcessor\"

echo.
echo Installation complete!
echo.
echo IMPORTANT: Edit config.ini to set your report template path
echo Location: %USERPROFILE%\GibbsCAMProcessor\config.ini
echo.
pause