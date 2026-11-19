@echo off
REM Builds dist\Quill\Quill.exe  (needs Python 3.10+ from python.org, "Add to PATH" ticked)
cd /d "%~dp0"
if not exist .venv-build (
    python -m venv .venv-build || goto :fail
)
call .venv-build\Scripts\activate.bat
python -m pip install --upgrade pip || goto :fail
python -m pip install -r requirements.txt pyinstaller || goto :fail
pyinstaller quill.spec --noconfirm --clean || goto :fail
echo.
echo Done:  dist\Quill\Quill.exe
echo Copy the whole dist\Quill folder to share it. Put extra theme .json files in dist\Quill\themes.
pause
exit /b 0
:fail
echo.
echo Build failed - scroll up for the error.
pause
exit /b 1
