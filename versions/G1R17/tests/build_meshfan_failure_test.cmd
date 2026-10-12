@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
ml64 /nologo /c /Fo meshfan_fp.obj ..\native\vba_fp.asm
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /EHsc /W4 /WX /std:c++17 meshfan_failure_injection.cpp meshfan_fp.obj /Fe:meshfan_failure_test.exe
if errorlevel 1 exit /b 1
meshfan_failure_test.exe
exit /b %errorlevel%
