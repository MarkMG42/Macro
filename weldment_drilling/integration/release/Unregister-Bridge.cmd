@echo off
setlocal
set "bridge_regasm=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
set "bridge_dll=%~dp0Macro.DrillingDimensionsBridge.dll"
if not exist "%bridge_regasm%" (
  echo ERROR: .NET Framework 4.x 64-bit RegAsm was not found.
  exit /b 1
)
if not exist "%bridge_dll%" (
  echo ERROR: Keep the bridge files in this folder until unregistration is complete.
  exit /b 1
)
"%bridge_regasm%" "%bridge_dll%" /unregister /nologo
set "bridge_result=%errorlevel%"
if not "%bridge_result%"=="0" (
  echo ERROR: Unregistration failed. Run this script as Administrator and check the message above.
  exit /b %bridge_result%
)
echo Unregistered Macro.DrillingDimensions. Restart SOLIDWORKS before deleting the folder.
exit /b 0
