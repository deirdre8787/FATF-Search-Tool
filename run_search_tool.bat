@echo off
title MER Search Tool v8
cd /d "%~dp0"

echo ==========================================================
echo   MER Search Tool  v8
echo   Search the 4th round, the 5th round, or both together
echo ==========================================================
echo.

rem ---- find a working Python -------------------------------------------
set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY if exist "%USERPROFILE%\AppData\Local\anaconda3\python.exe" set "PY=%USERPROFILE%\AppData\Local\anaconda3\python.exe"
if not defined PY if exist "C:\Users\LUO_D\AppData\Local\anaconda3\python.exe" set "PY=C:\Users\LUO_D\AppData\Local\anaconda3\python.exe"
if not defined PY if exist "%USERPROFILE%\anaconda3\python.exe" set "PY=%USERPROFILE%\anaconda3\python.exe"
if not defined PY if exist "C:\ProgramData\anaconda3\python.exe" set "PY=C:\ProgramData\anaconda3\python.exe"

if not defined PY (
  echo Python was not found on this machine.
  echo Open Anaconda Prompt and run:   python "%~dp0app.py"
  echo.
  pause
  exit /b 1
)

rem ---- make sure flask is available (rapidfuzz is optional) ------------
"%PY%" -c "import flask" >nul 2>&1
if errorlevel 1 (
  echo First run: installing Flask, please wait...
  "%PY%" -m pip install flask rapidfuzz --quiet
  echo.
)

echo Starting... your browser will open automatically.
echo KEEP THIS WINDOW OPEN while using the tool. Close it to stop.
echo.
"%PY%" app.py

if errorlevel 1 (
  echo.
  echo The tool stopped with an error. Read the message above.
  echo Common causes: port 5177 already in use ^(close any other copy of this
  echo window^), or a source CSV is missing from a round folder.
  echo.
  pause
)
