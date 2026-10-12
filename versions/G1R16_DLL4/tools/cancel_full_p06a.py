"""User-authorized cancellation of this task's own full P06A process only."""
import json,time
import win32com.client,win32process
from prepare import ROOT,write
def main():
 s=json.loads((ROOT/'results/p06a_full_dll4/status.json').read_text(encoding='utf-8'));assert s['status']=='running'
 app=win32com.client.GetActiveObject('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1];assert pid==s['excel_pid'],(pid,s['excel_pid'])
 book=app.Workbooks('P06A_dll4.xlsm');assert str(book.FullName).lower()==str(ROOT/'results/p06a_full_dll4/P06A_dll4.xlsm').lower()
 record={'user_requested_stop':True,'excel_pid':pid,'method':'existing RPX_Stop macro; no process kill','requested_at':time.strftime('%Y-%m-%d %H:%M:%S')}
 write(ROOT/'results/p06a_full_dll4/cancellation.json',record)
 app.Run("'"+book.Name+"'!RPX_Stop");record['stop_macro_returned']=True;write(ROOT/'results/p06a_full_dll4/cancellation.json',record);print('RPX_Stop requested',pid,flush=True)
if __name__=='__main__':main()
