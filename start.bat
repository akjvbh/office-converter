@echo off
rem Keep this file ASCII-only with CRLF line endings: cmd mis-parses batch
rem files that combine non-ASCII text or goto-inside-if with LF-only newlines.
rem Chinese messages and logging live in launcher.py.
cd /d "%~dp0"

if exist ".venv-ocr\Scripts\python.exe" goto usevenv

where py >nul 2>nul
if errorlevel 1 goto nopython

echo Virtualenv not found at .venv-ocr\Scripts\python.exe
echo Falling back to the "py" launcher.
echo.
py -3 launcher.py
goto done

:usevenv
".venv-ocr\Scripts\python.exe" launcher.py
goto done

:nopython
echo ERROR: no usable Python found.
echo.
echo   Missing: .venv-ocr\Scripts\python.exe
echo   Also missing: the "py" launcher
echo.
echo Do this once, then run start.bat again:
echo.
echo   py -3 -m venv .venv-ocr
echo   .venv-ocr\Scripts\python.exe -m pip install -r requirements.txt
echo.
echo Note: typing "python" may open the Microsoft Store instead of Python.
echo See README.md, section "Install", for details.

:done
echo.
echo Service stopped. Log file: _start_log.txt
pause
