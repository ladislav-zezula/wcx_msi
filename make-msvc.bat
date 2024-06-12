:: Build file for Visual Studio 2008 and 2017
@echo off

:: Save the values of INCLUDE, LIB and PATH
set PROJECT_DIR=%~dp0
set SAVE_INCLUDE=%INCLUDE%
set SAVE_PATH=%PATH%
set SAVE_LIB=%LIB%
set PROJECT_NAME=wcx_msi
set PROJECT_ZIP_NAME=wcx_msi

:: Remember whether we shall publish the project
if "x%1" == "x/web" set PUBLISH_PROJECT=1

:: Determine where the program files are, both for 64-bit and 32-bit Windows
if exist "%ProgramW6432%"      set PROGRAM_FILES_X64=%ProgramW6432%
if exist "%ProgramFiles%"      set PROGRAM_FILES_DIR=%ProgramFiles%
if exist "%ProgramFiles(x86)%" set PROGRAM_FILES_DIR=%ProgramFiles(x86)%

:: Determine the installed version of Visual Studio (Prioritize Enterprise over Professional)
if exist "%PROGRAM_FILES_DIR%\Microsoft Visual Studio 9.0\VC\vcvarsall.bat"                               set VCVARS_2008=%PROGRAM_FILES_DIR%\Microsoft Visual Studio 9.0\VC\vcvarsall.bat
if exist "%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2017\Enterprise\VC\Auxiliary\Build\vcvarsall.bat"   set VCVARS_20xx=%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2017\Enterprise\VC\Auxiliary\Build\vcvarsall.bat
if exist "%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2017\Professional\VC\Auxiliary\Build\vcvarsall.bat" set VCVARS_20xx=%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2017\Professional\VC\Auxiliary\Build\vcvarsall.bat
if exist "%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2017\Community\VC\Auxiliary\Build\vcvarsall.bat"    set VCVARS_20xx=%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2017\Community\VC\Auxiliary\Build\vcvarsall.bat
if exist "%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2019\Enterprise\VC\Auxiliary\Build\vcvarsall.bat"   set VCVARS_20xx=%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2019\Enterprise\VC\Auxiliary\Build\vcvarsall.bat
if exist "%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2019\Professional\VC\Auxiliary\Build\vcvarsall.bat" set VCVARS_20xx=%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2019\Professional\VC\Auxiliary\Build\vcvarsall.bat
if exist "%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat"    set VCVARS_20xx=%PROGRAM_FILES_DIR%\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat
if exist "%PROGRAM_FILES_X64%\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvarsall.bat"   set VCVARS_20xx=%PROGRAM_FILES_X64%\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvarsall.bat
if exist "%PROGRAM_FILES_X64%\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvarsall.bat" set VCVARS_20xx=%PROGRAM_FILES_X64%\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvarsall.bat
if exist "%PROGRAM_FILES_X64%\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat"    set VCVARS_20xx=%PROGRAM_FILES_X64%\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat

:: Build the project using Visual Studio 2008 and 2017
rem if not "x%VCVARS_2008%" == "x" call :BuildProject "%VCVARS_2008%" %PROJECT_NAME%_vs08.sln En en x86 Win32
rem if not "x%VCVARS_2008%" == "x" call :BuildProject "%VCVARS_2008%" %PROJECT_NAME%_vs08.sln En en x64 x64
if not "x%VCVARS_20xx%" == "x" call :BuildProject "%VCVARS_20xx%" %PROJECT_NAME%.sln x64 x64
if not "x%VCVARS_20xx%" == "x" call :BuildProject "%VCVARS_20xx%" %PROJECT_NAME%.sln x86 Win32
PostBuild.exe %PROJECT_NAME%.rc

:: Publish the project
if not "x%PUBLISH_PROJECT%" == "x1" goto:eof
echo [*] Publishing project ...
call make-zip.bat .\bin\x64\Release\msi.wcx64 .\bin\Win32\Release\msi.wcx
goto:eof

::-----------------------------------------------------------------------------
:: Build the project
::
:: Parameters:
::
::   %1     Full path to the VCVARS.BAT file
::   %2     Plain name of the .sln solution file
::   %3     x86 or x64
::   %4     Win32 or x64
::

:BuildProject
echo [*] Building %PROJECT_NAME% (%3) ...
call %1 %3 >nul
devenv.com %2 /project "%PROJECT_NAME%" /rebuild "Release|%4" >nul

:: Restore environment variables to the old level
:RestoreEnvironment
set INCLUDE=%SAVE_INCLUDE%
set PATH=%SAVE_PATH%
set LIB=%SAVE_LIB%

:: Delete environment variables that are set by Visual Studio
set __VSCMD_PREINIT_PATH=
set EXTERNAL_INCLUDE=
set VSINSTALLDIR=
set VCINSTALLDIR=
set DevEnvDir=
set LIBPATH=

