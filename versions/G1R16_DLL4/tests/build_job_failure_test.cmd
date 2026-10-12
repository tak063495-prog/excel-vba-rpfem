@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
ml64 /nologo /c /Fo job_fp.obj ..\native\vba_fp.asm
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 /Fojob_native.obj ..\native\rpfem_native.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /EHsc /W4 /WX /std:c++17 job_failure_injection.cpp job_fp.obj job_native.obj /Fe:refine_failure_injection.exe
if errorlevel 1 exit /b 1
refine_failure_injection.exe
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /EHsc /W4 /WX /std:c++17 /DTEST_HLINEAR job_failure_injection.cpp ..\native\rpfem_lu.cpp job_fp.obj job_native.obj /Fe:hlinear_failure_injection.exe
if errorlevel 1 exit /b 1
hlinear_failure_injection.exe
exit /b %errorlevel%
