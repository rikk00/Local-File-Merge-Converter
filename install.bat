@echo off
setlocal
for %%I in ("%~dp0.") do set "INSTALL_DIR=%%~fI"
set "INSTALL_SCRIPT=%~f0"
set "MERGE_SCRIPT=%INSTALL_DIR%\merge_pdf.py"

fltmc >nul 2>&1
if errorlevel 1 (
    rem Elevation is required to write the HKEY_CLASSES_ROOT shell entries.
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$script = $env:INSTALL_SCRIPT; Start-Process -FilePath $script -ArgumentList 'elevated' -WorkingDirectory $env:INSTALL_DIR -Verb RunAs"
    if errorlevel 1 (
        echo Administrator elevation was cancelled or failed.
        pause
        exit /b 1
    )
    exit /b 0
)

for /f "delims=" %%P in ('where pythonw 2^>nul') do call :select_python "%%P"
if not defined PYTHONW (
    for /f "delims=" %%P in ('py -3 -c "import sys; from pathlib import Path; print(Path(sys.executable).with_name('pythonw.exe'))" 2^>nul') do call :select_python "%%P"
)
if not defined PYTHONW (
    echo Python with pythonw.exe was not found. Install Python or add it to PATH.
    pause
    exit /b 1
)

"%PYTHON%" -m pip install -r "%INSTALL_DIR%\requirements.txt"
if errorlevel 1 goto :failed

reg add "HKCR\*\shell\ConvertTo" /v MUIVerb /t REG_SZ /d "Convert to" /f
if errorlevel 1 goto :failed
reg add "HKCR\*\shell\ConvertTo" /v SubCommands /t REG_SZ /d "" /f
if errorlevel 1 goto :failed
reg add "HKCR\*\shell\ConvertTo\shell\Markdown" /v MUIVerb /t REG_SZ /d "Markdown" /f
if errorlevel 1 goto :failed
reg add "HKCR\*\shell\ConvertTo\shell\Markdown\command" /ve /t REG_SZ /d "\"%PYTHONW%\" \"%INSTALL_DIR%\converter.py\" \"%%1\" --target md" /f
if errorlevel 1 goto :failed
reg add "HKCR\*\shell\ConvertTo\shell\PDF" /v MUIVerb /t REG_SZ /d "PDF" /f
if errorlevel 1 goto :failed
reg add "HKCR\*\shell\ConvertTo\shell\PDF\command" /ve /t REG_SZ /d "\"%PYTHONW%\" \"%INSTALL_DIR%\converter.py\" \"%%1\" --target pdf" /f
if errorlevel 1 goto :failed

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$shell = New-Object -ComObject WScript.Shell; $shortcut = $shell.CreateShortcut((Join-Path $env:APPDATA 'Microsoft\Windows\SendTo\Merge as PDF.lnk')); $shortcut.TargetPath = $env:PYTHONW; $shortcut.Arguments = ('"' + $env:MERGE_SCRIPT + '"'); $shortcut.WorkingDirectory = $env:INSTALL_DIR; $shortcut.Save()"
if errorlevel 1 goto :failed

echo Installation completed successfully.
pause
exit /b 0

:select_python
if not defined PYTHONW if exist "%~dpn1.exe" (
    set "PYTHONW=%~f1"
    set "PYTHON=%~dpn1.exe"
)
exit /b 0

:failed
echo Installation failed. Review the error above and try again.
pause
exit /b 1
