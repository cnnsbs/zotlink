@echo off
setlocal
rem ============================================================
rem  Zotero citation linker - drag .docx file(s) onto this file
rem  (or onto its shortcut). Double-click = file selection dialog.
rem  Keep this .bat in the SAME folder as zotero_link_citations.py
rem ============================================================

set "ENV_NAME=zotlink"
set "SCRIPT=%~dp0zotero_link_citations.py"

rem --- find Anaconda / Miniconda
set "CONDA_ROOT="
for %%D in ("%USERPROFILE%\anaconda3" "%USERPROFILE%\miniconda3" "C:\ProgramData\anaconda3" "C:\ProgramData\miniconda3" "%LOCALAPPDATA%\anaconda3" "%LOCALAPPDATA%\miniconda3" "C:\anaconda3" "D:\anaconda3") do (
    if not defined CONDA_ROOT if exist "%%~D\Scripts\activate.bat" set "CONDA_ROOT=%%~D"
)
if not defined CONDA_ROOT (
    echo [ERROR] Anaconda folder not found. Edit CONDA_ROOT in this .bat file.
    pause & exit /b 1
)
if not exist "%SCRIPT%" (
    echo [ERROR] zotero_link_citations.py not found next to this .bat:
    echo         %SCRIPT%
    pause & exit /b 1
)

call "%CONDA_ROOT%\Scripts\activate.bat" %ENV_NAME%
if errorlevel 1 (
    echo [ERROR] conda env "%ENV_NAME%" could not be activated.
    pause & exit /b 1
)

if "%~1"=="" (
    python "%SCRIPT%"
    goto done
)

:loop
if "%~1"=="" goto done
set "NAME=%~n1"
if /i not "%~x1"==".docx" (
    echo [SKIP] not a .docx file: %~nx1
    goto next
)
if /i "%NAME:~-7%"=="_zotero" (
    echo [SKIP] already converted: %~nx1
    goto next
)
echo.
echo ================= %~nx1 =================
python "%SCRIPT%" "%~f1"
:next
shift
goto loop

:done
echo.
pause
