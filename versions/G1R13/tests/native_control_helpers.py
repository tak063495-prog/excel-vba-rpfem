from pathlib import Path
import sys,json,shutil,gc
import win32com.client
from native_helpers import compile_book
ROOT=Path(__file__).resolve().parents[1]
def write(p,obj):p.write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding='utf-8')
def call(app,book,name,*args):return app.Run("'"+book.Name+"'!"+name,*args)
def inject(book,name,text):
    try:c=book.VBProject.VBComponents(name);c.CodeModule.DeleteLines(1,c.CodeModule.CountOfLines)
    except Exception:c=book.VBProject.VBComponents.Add(1);c.Name=name
    c.CodeModule.AddFromString('\r\n'.join(row for row in text.splitlines() if not row.startswith('Attribute VB_')))
