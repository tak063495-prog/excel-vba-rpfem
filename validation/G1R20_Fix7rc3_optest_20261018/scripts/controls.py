"""Acknowledge only own Excel OK dialogs; cooperative STOP on user request."""
import threading,time,traceback
import pythoncom,win32com.client,win32gui,win32process,win32con
from prepare import ROOT,write
def start_controls(app,pid,bookname,folder):
    stop=threading.Event();dialogs=[]
    stream=pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch,app._oleobj_)
    def monitor():
        pythoncom.CoInitialize()
        proxy=win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(stream,pythoncom.IID_IDispatch))
        cancellation_sent=False
        try:
            while not stop.wait(.5):
                def window(h,_):
                    if win32process.GetWindowThreadProcessId(h)[1]!=pid or not win32gui.IsWindowVisible(h) or win32gui.GetClassName(h)!='#32770':return
                    texts=[];buttons=[]
                    def child(c,_):
                        text=win32gui.GetWindowText(c)
                        if text:texts.append(text)
                        if text=='OK':buttons.append(c)
                    win32gui.EnumChildWindows(h,child,None)
                    if buttons:
                        dialogs.append({'time':time.time(),'text':texts});write(folder/'dialogs.json',dialogs)
                        for c in buttons:win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
                win32gui.EnumWindows(window,None)
                if ((ROOT/'STOP').exists() or (folder/'STOP_CASE').exists()) and not cancellation_sent:
                    try:
                        proxy.Run("'"+bookname+"'!RPX_Stop");cancellation_sent=True
                        write(folder/'cooperative_stop.json',{'sent':True,'time':time.time(),'pid':pid,'reason':'TIMEOUT' if (folder/'STOP_CASE').exists() else 'USER_STOP'})
                    except pythoncom.com_error as e:
                        write(folder/'cooperative_stop_attempt.json',{'sent':False,'error':repr(e),'time':time.time()})
        except Exception:
            write(folder/'control_error.json',{'error':traceback.format_exc()})
        finally:
            proxy=None;pythoncom.CoUninitialize()
    threading.Thread(target=monitor,daemon=True).start()
    return stop
