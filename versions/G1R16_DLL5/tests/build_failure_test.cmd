@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
cl /nologo /O2 /MT /EHsc /W4 /WX /std:c++17 order_failure_injection.cpp /Fe:order_failure_injection.exe
if errorlevel 1 exit /b 1
order_failure_injection.exe
exit /b %errorlevel%
