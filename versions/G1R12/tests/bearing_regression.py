from pathlib import Path
import shutil,json
import win32com.client
from native_regression import ROOT,setting
from build_and_workflow import OUTPUT
from native_helpers import compile_book
def run():
    path=ROOT/'tests/bearing_G1R12.xlsm';shutil.copy2(ROOT/OUTPUT,path)
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
    book=None;result={'scope':'small native nonadaptive/adaptive bearing workflow; not full bearing/rigid regression'}
    try:
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
        compile_book(app,book)
        result['nonadaptive']=call('RPX_TestBearingWorkflow');assert result['nonadaptive'].startswith('PASS'),result
        for key,val in [('ADAPT',1),('TARGET',16),('CYCLES',0),('QUADRATURE',2),('C0_AUTO',0)]:setting(book,key,val)
        call('RPX_Run')
        ws=book.Worksheets('解析結果')
        result['adaptive']={'status':ws.Range('A1').Value,'multipliers':list(ws.Range('B4:C4').Value[0]),'pressures':list(ws.Range('B5:C5').Value[0])}
        assert all(abs(v-3.5)<1e-5 for v in result['adaptive']['multipliers']),result
        assert all(abs(v-2)<1e-5 for v in result['adaptive']['pressures']),result
        result['pass']=True;book.Save()
        print('Native bearing regression PASS',result,flush=True)
    finally:
        (ROOT/'results/native_bearing.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        if book is not None:book.Close(SaveChanges=False)
        app.Quit()
if __name__=='__main__':run()
