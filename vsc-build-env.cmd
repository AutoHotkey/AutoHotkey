@echo off

:: Assume build environment is already setup if msbuild can be located
where msbuild >nul 2>nul && exit /b 0

:: Allow the path to vsdevcmd to be provided by our caller
if exist "%vsdevcmd%" "%vsdevcmd%"
:: If we're still running, must be no vsdevcmd

if "%ProgramFiles(x86)%"=="" set ProgramFiles(x86)=%ProgramFiles%
set vswhere="%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist %vswhere% (
    echo vswhere.exe not found; unable to locate build tools.
    exit 1
)

:: This should work for Visual Studio
for /f "usebackq delims=" %%i in (`%vswhere% -latest -requires Microsoft.VisualStudio.Workload.NativeDesktop -find *\Tools\vsdevcmd.bat`) do "%%i"
:: This should work with Visual Studio Build Tools
for /f "usebackq delims=" %%i in (`%vswhere% -latest -products * -requires Microsoft.VisualStudio.Workload.VCTools -find *\Tools\vsdevcmd.bat`) do "%%i"
:: As a last resort, try without specifying the required workload
for /f "usebackq delims=" %%i in (`%vswhere% -latest -products * -find *\Tools\vsdevcmd.bat`) do "%%i"
:: If we're still running, vsdevcmd wasn't executed
echo Unable to locate build tools.
exit 1