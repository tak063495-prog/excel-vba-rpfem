import gc,json,shutil,argparse
import win32com.client
from prepare import ROOT,write,sha
from test_hsd_regression import inject
from native_helpers import compile_book

def run():
 p=argparse.ArgumentParser();p.add_argument('--tag',required=True);p.add_argument('--expect-failure',action='store_true');args=p.parse_args();assert args.tag.isidentifier()
 folder=ROOT/'results'/args.tag;folder.mkdir(parents=True,exist_ok=False)
 path=folder/'ModeSwitch.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm',path);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0);inject(book,'ContractTests');compile_book(app,book)
  msg=str(app.Run("'"+book.Name+"'!TestLUModeSwitch"));print(msg,flush=True)
  record={'message':msg,'book_sha256':sha(path),'dll_sha256':sha(folder/'rpfem_native_x64.dll'),'pass':msg.startswith('PASS')};write(folder/'summary.json',record)
  assert record['pass'] != args.expect_failure
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
if __name__=='__main__':run()
