@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
ml64 /nologo /c /Fo dll5_fp.obj ..\native\vba_fp.asm
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 /Fodll5_native.obj ..\native\rpfem_native.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /EHsc /W4 /WX /std:c++17 /DTEST_COMPILE dll5_failure_injection.cpp ..\native\rpfem_order.cpp dll5_fp.obj /Fe:compile_failure_test.exe
if errorlevel 1 exit /b 1
compile_failure_test.exe
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /EHsc /W4 /WX /std:c++17 /DTEST_SCHUR dll5_failure_injection.cpp dll5_fp.obj dll5_native.obj /Fe:schur_failure_test.exe
if errorlevel 1 exit /b 1
schur_failure_test.exe
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /EHsc /W4 /WX /std:c++17 dll5_failure_injection.cpp ..\native\rpfem_lu.cpp dll5_fp.obj dll5_native.obj /Fe:krylov_failure_test.exe
if errorlevel 1 exit /b 1
krylov_failure_test.exe
exit /b %errorlevel%
