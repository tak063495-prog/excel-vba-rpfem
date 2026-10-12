@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
ml64 /nologo /c /Fo lu_test_fp.obj ..\native\vba_fp.asm
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /EHsc /W4 /WX /std:c++17 lu_failure_injection.cpp lu_test_fp.obj /Fe:lu_failure_injection.exe
if errorlevel 1 exit /b 1
lu_failure_injection.exe
exit /b %errorlevel%
