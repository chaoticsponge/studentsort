@echo off
setlocal
cd /d "%~dp0"

set "LOG=%~dp0build_log.txt"
echo Build started > "%LOG%"
echo Folder: %CD% >> "%LOG%"

set "PYTHON_CMD="

python --version >> "%LOG%" 2>&1
if not errorlevel 1 set "PYTHON_CMD=python"

if "%PYTHON_CMD%"=="" (
  python3 --version >> "%LOG%" 2>&1
  if not errorlevel 1 set "PYTHON_CMD=python3"
)

if "%PYTHON_CMD%"=="" (
  py -3 --version >> "%LOG%" 2>&1
  if not errorlevel 1 set "PYTHON_CMD=py -3"
)

if "%PYTHON_CMD%"=="" (
  echo No Windows Python command was found. >> "%LOG%"
  echo No Windows Python command was found.
  echo Install Python for Windows inside this Whisky bottle, then run this file again.
  echo Use the Windows installer from python.org and tick "Add python.exe to PATH".
  pause
  exit /b 1
)

echo Using Python command: %PYTHON_CMD% >> "%LOG%"
%PYTHON_CMD% -m pip install --upgrade pyinstaller >> "%LOG%" 2>&1
if errorlevel 1 (
  echo PyInstaller install failed. See build_log.txt.
  pause
  exit /b 1
)

%PYTHON_CMD% -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name "ExamAccessLogGenerator" ^
  --add-data "Exam Room Incident Log Template.docx;." ^
  exam_access_app.py >> "%LOG%" 2>&1

if errorlevel 1 (
  echo EXE build failed. See build_log.txt.
  pause
  exit /b 1
)

if not exist "dist\ExamAccessLogGenerator.exe" (
  echo Build finished but dist\ExamAccessLogGenerator.exe was not found. >> "%LOG%"
  echo Build finished but the EXE was not found. See build_log.txt.
  pause
  exit /b 1
)

copy /Y "dist\ExamAccessLogGenerator.exe" "%~dp0ExamAccessLogGenerator.exe" >> "%LOG%" 2>&1
if exist "%USERPROFILE%\Desktop" (
  copy /Y "dist\ExamAccessLogGenerator.exe" "%USERPROFILE%\Desktop\ExamAccessLogGenerator.exe" >> "%LOG%" 2>&1
)
if exist "C:\Users\Public\Desktop" (
  copy /Y "dist\ExamAccessLogGenerator.exe" "C:\Users\Public\Desktop\ExamAccessLogGenerator.exe" >> "%LOG%" 2>&1
)

echo.
echo Done. The Windows EXE is here:
echo %~dp0ExamAccessLogGenerator.exe
echo.
echo If Desktop copying worked, another copy is also on the Windows Desktop.
echo A build log is here:
echo %LOG%
explorer.exe /select,"%~dp0ExamAccessLogGenerator.exe"
pause
