@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
if not exist ..\delivery mkdir ..\delivery
ml64 /nologo /c vba_fp.asm
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 rpfem_native.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /LD /EHsc /W4 /WX /std:c++17 rpfem_order.cpp rpfem_lu.cpp rpfem_refine.cpp rpfem_hlinear.cpp rpfem_guard.cpp rpfem_batch.cpp rpfem_schur.cpp rpfem_krylov.cpp rpfem_compile.cpp rpfem_native.obj vba_fp.obj /Fe:..\delivery\rpfem_native_x64.dll /link /INCREMENTAL:NO oleaut32.lib
exit /b %errorlevel%
