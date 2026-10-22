@echo off
cd /d C:\Users\link_\rpfem_bench\repo\validation\G1R20_Fix9rc9b_p06n_20261021
set PYTHONIOENCODING=utf-8
"C:\Users\link_\rpfem_bench\.venv\Scripts\python.exe" -X utf8 scripts\runner.py batch >> results\batch_stdout.txt 2>&1
echo EXIT %ERRORLEVEL% %DATE% %TIME% >> results\batch_stdout.txt

