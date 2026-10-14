@echo off
rem Windows launcher for Word Formatter.
rem
rem IMPORTANT: this file is intentionally ASCII-only, comments included.
rem cmd.exe parses batch files using the current code page and mangles
rem non-ASCII characters, so parts of Russian words get executed as commands.
rem All Russian text is printed from tools\messages\*.txt as data.
rem
rem The job of this file is to find a suitable Python and hand over to
rem tools\bootstrap.py: it checks the version, creates the local .venv
rem environment, installs dependencies and starts the application.
rem
rem Administrator rights are not needed and nothing is installed system-wide.

setlocal EnableExtensions
pushd "%~dp0"

if not exist "tools\bootstrap.py" goto :incomplete
if not exist "tools\messages\no_python.txt" goto :incomplete

set "PYTHON_CMD="

rem 1. The official Python launcher, installed together with python.org builds.
py -3 -c "print('ok')" >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3"

rem 2. The python command from PATH. The output check is needed to tell a real
rem    Python from the Microsoft Store stub: the stub opens the Store app and
rem    prints nothing at all.
if defined PYTHON_CMD goto :found
python -c "print('ok')" >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=python"

:found
if not defined PYTHON_CMD goto :no_python

%PYTHON_CMD% "tools\bootstrap.py"
set "EXIT_CODE=%ERRORLEVEL%"
popd
if not "%EXIT_CODE%"=="0" goto :failed
exit /b 0

:no_python
chcp 65001 >nul
type "tools\messages\no_python.txt"
echo.
popd
pause
exit /b 2

:incomplete
chcp 65001 >nul
type "tools\messages\incomplete_copy.txt"
echo.
popd
pause
exit /b 1

:failed
chcp 65001 >nul
type "tools\messages\failed.txt"
echo.
echo Exit code: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
