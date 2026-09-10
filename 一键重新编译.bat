@echo off
cd /d "%~dp0"
set "CSCDIR=C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
if not exist "%CSCDIR%\csc.exe" set "CSCDIR=C:\Windows\Microsoft.NET\Framework\v4.0.30319"

echo Compiling FinancePdfTool.cs ...
"%CSCDIR%\csc.exe" /nologo /target:winexe /optimize+ /win32icon:"%~dp0app.ico" /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:"%CSCDIR%\System.Runtime.dll" /r:"%CSCDIR%\System.Runtime.WindowsRuntime.dll" /r:"C:\Windows\System32\WinMetadata\Windows.Foundation.winmd" /r:"C:\Windows\System32\WinMetadata\Windows.Data.winmd" /r:"C:\Windows\System32\WinMetadata\Windows.Storage.winmd" /out:"%~dp0财务专用PDF转换器.exe" "%~dp0FinancePdfTool.cs"

if %ERRORLEVEL% equ 0 (
    echo.
    echo ===================================================
    echo [SUCCESS] Build completed: FinancePdfTool.exe
    echo ===================================================
) else (
    echo.
    echo [ERROR] Build failed! Check errors above.
)
echo.
pause
