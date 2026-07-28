@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "CSC="
if exist "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set "CSC=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not defined CSC if exist "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe" set "CSC=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if not defined CSC (
  echo [ERROR] csc.exe not found
  exit /b 1
)

if not exist "bin\Release" mkdir "bin\Release"

echo Using: %CSC%
"%CSC%" /nologo /target:winexe /platform:anycpu /optimize+ /out:bin\Release\OutlookEmlExport.exe /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Xml.dll /reference:System.Xml.Linq.dll /reference:Microsoft.CSharp.dll /reference:System.Web.Extensions.dll Program.cs MainForm.cs MainForm.Designer.cs Export\Utils.cs Export\ExportState.cs Export\EmlBuilder.cs Export\OutlookExporter.cs Properties\AssemblyInfo.cs
if errorlevel 1 (
  echo [ERROR] build failed
  exit /b 1
)

copy /Y App.config bin\Release\OutlookEmlExport.exe.config >nul
echo.
echo OK: bin\Release\OutlookEmlExport.exe
endlocal
