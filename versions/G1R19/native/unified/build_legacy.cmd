@echo off
setlocal
call "%~1" >nul
if errorlevel 1 exit /b 1
pushd "%~dp0obj\legacy"
ml64 /nologo /c /Fo vba_fp.obj "%~dp0legacy\vba_fp.asm"
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 "%~dp0legacy\rpfem_native.c"
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /EHsc /W4 /WX /std:c++17 "%~dp0legacy\rpfem_order.cpp" "%~dp0legacy\rpfem_lu.cpp" "%~dp0legacy\rpfem_refine.cpp" "%~dp0legacy\rpfem_hlinear.cpp" "%~dp0legacy\rpfem_guard.cpp" "%~dp0legacy\rpfem_batch.cpp" "%~dp0legacy\rpfem_schur.cpp" "%~dp0legacy\rpfem_krylov.cpp" "%~dp0legacy\rpfem_compile.cpp" "%~dp0legacy\rpfem_meshfan.cpp"
if errorlevel 1 exit /b 1
link /NOLOGO /DLL /INCREMENTAL:NO /Brepro /OUT:"%~dp0bin\rpfem_native_x64.dll" /IMPLIB:"%~dp0bin\rpfem_native_x64.lib" /DEF:"%~dp0accuracy_exports.def" *.obj "%~dp0accuracy\target\x86_64-pc-windows-msvc\release\rpfem_accuracy_x64.lib" @"%~dp0obj\rust_native_static_libs.rsp" oleaut32.lib
if errorlevel 1 exit /b 1
dumpbin /NOLOGO /EXPORTS "%~dp0bin\rpfem_native_x64.dll" > "%~dp0bin\exports.txt"
if errorlevel 1 exit /b 1
dumpbin /NOLOGO /DEPENDENTS "%~dp0bin\rpfem_native_x64.dll" > "%~dp0bin\dependencies.txt"
exit /b %errorlevel%
