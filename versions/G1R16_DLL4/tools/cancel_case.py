"""Stop only the identified task-owned test via its normal cancellation macro."""
import argparse,json
import win32com.client,win32process
from prepare import ROOT,write
def main():
 p=argparse.ArgumentParser();p.add_argument('--tag',required=True);p.add_argument('--pid',type=int,required=True);p.add_argument('--case',required=True);p.add_argument('--backend',required=True);p.add_argument('--reason',required=True);args=p.parse_args()
 assert args.tag.isidentifier() and args.case in ['P06A','P06N','D03N','D07A'] and args.backend in ['dll3','dll4']
 folder=ROOT/'results'/args.tag;s=json.loads((folder/'status.json').read_text(encoding='utf-8'));assert s['status']=='running' and s['excel_pid']==args.pid
 app=win32com.client.GetActiveObject('Excel.Application');assert win32process.GetWindowThreadProcessId(app.Hwnd)[1]==args.pid
 book=app.Workbooks(args.case+'_'+args.backend+'.xlsm');assert str(book.FullName).lower()==str(folder/(args.case+'_'+args.backend+'.xlsm')).lower()
 r={'reason':args.reason,'user_requested_stop':False,'agent_bounded_validation':True,'excel_pid':args.pid,'workbook':str(book.FullName),'stop_macro_returned':False,'process_killed':False}
 write(folder/'cancellation.json',r);app.Run("'"+book.Name+"'!RPX_Stop");r['stop_macro_returned']=True;write(folder/'cancellation.json',r)
 print('Normal stop returned for owned test',args.tag,flush=True)
if __name__=='__main__':main()
