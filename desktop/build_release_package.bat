@echo off
setlocal
cd /d "%~dp0"
if not exist dist\SORTING_OS.exe (
    echo.
    echo ERROR: Please build the EXE first by running build_windows.bat on a Windows target.
    echo.
    exit /b 1
)
if not exist release mkdir release
powershell -NoProfile -Command "Compress-Archive -Path '.\dist\SORTING_OS.exe', '.\README_FA.md', '.\RELEASE_READY_FA.md', '.\requirements.txt' -DestinationPath '.\release\SORTING_OS_Final_Package.zip' -Force"
echo.
echo Release package created at: %~dp0release\SORTING_OS_Final_Package.zip
echo.
