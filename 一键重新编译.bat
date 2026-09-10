@echo off
cd /d "%~dp0
set CSCDIR=C:\Windows\Microsoft.NET\Framework64\v4.0.30319
if not exist %CSCDIR%\csc.exe set CSCDIR=C:\Windows\Microsoft.NET\Framework\v4.0.30319
echo Compiling FinancePdfTool.exe...
%CSCDIR%\csc.exe /nologo /target:winexe /optimize+ /platform:anycpu /win32icon:%~dp0app.ico /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /out:%~dp0财务专用PDF转换器.exe %~dp0FinancePdfTool.cs
if %ERRORLEVEL% equ 0 (
 echo [SUCCESS] Build completed successfully!
) else (
 echo [ERROR] Build failed!
)
pause
