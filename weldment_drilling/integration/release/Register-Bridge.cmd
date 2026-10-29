@echo off
setlocal
set "bridge_regasm=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
set "bridge_dll=%~dp0Macro.DrillingDimensionsBridge.dll"
if not exist "%bridge_regasm%" (
  echo ERROR: .NET Framework 4.x 64-bit RegAsm was not found.
  echo Install/enable .NET Framework 4.8 and retry.
  exit /b 1
)
if not exist "%bridge_dll%" (
  echo ERROR: Macro.DrillingDimensionsBridge.dll must be beside this script.
  exit /b 1
)
if not exist "%~dp0SolidWorks.Interop.sldworks.dll" (
  echo ERROR: SolidWorks.Interop.sldworks.dll is missing beside the bridge.
  exit /b 1
)
if not exist "%~dp0SolidWorks.Interop.swconst.dll" (
  echo ERROR: SolidWorks.Interop.swconst.dll is missing beside the bridge.
  exit /b 1
)
echo Registering the 64-bit drilling dimensions bridge from this permanent folder.
echo Administrator rights are required. No original AutoSizeDrilling DLL is registered.
"%bridge_regasm%" "%bridge_dll%" /codebase /nologo
set "bridge_result=%errorlevel%"
if not "%bridge_result%"=="0" (
  echo ERROR: Registration failed. Run this script as Administrator and check the message above.
  exit /b %bridge_result%
)
set "bridge_cscript=%SystemRoot%\System32\cscript.exe"
if exist "%SystemRoot%\Sysnative\cscript.exe" set "bridge_cscript=%SystemRoot%\Sysnative\cscript.exe"
"%bridge_cscript%" //nologo "%~dp0Check-Bridge.vbs"
if errorlevel 1 (
  echo ERROR: Registration completed but activation check failed. Keep this folder and inspect the error.
  exit /b 1
)
echo Registered and activated Macro.DrillingDimensions. Restart SOLIDWORKS before using it.
exit /b 0
