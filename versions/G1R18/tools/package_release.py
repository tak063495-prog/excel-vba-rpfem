"""Build, reopen and hash the isolated G1R18 workbook and full distribution."""
from pathlib import Path
import gc
import hashlib
import json
import shutil
import time
import zipfile

import psutil
import win32com.client
import win32process

ROOT=Path(__file__).resolve().parents[1]
BASE=ROOT.parent/'G1R17'
DELIVERY=ROOT/'delivery'
NAME='RPFEM_20261014_G1R18'


def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()


def text(path):
    return '\r\n'.join(line for line in path.read_text(encoding='utf-8').splitlines() if not line.startswith('Attribute VB_'))


def compile_book(app,book):
    book.Activate();book.VBProject.VBComponents('RPX_AccuracyNative').CodeModule.CodePane.Show()
    control=app.VBE.CommandBars.FindControl(1,578)
    if control.Enabled: control.Execute();time.sleep(.3)
    if control.Enabled: raise RuntimeError('Distribution VBA compilation failed')
    app.VBE.MainWindow.Visible=False
    return True


def create_book():
    DELIVERY.mkdir(parents=True,exist_ok=True)
    path=DELIVERY/(NAME+'.xlsm')
    shutil.copyfile(BASE/'delivery/RPFEM_20261012_G1R17_P06N_Coarse_Checked.xlsm',path)
    shutil.copyfile(BASE/'delivery/rpfem_native_x64.dll',DELIVERY/'rpfem_native_x64.dll')
    shutil.copyfile(ROOT/'native/accuracy/bin/rpfem_accuracy_x64.dll',DELIVERY/'rpfem_accuracy_x64.dll')
    prior={p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE'}
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    assert pid not in prior
    report=dict(state='RUNNING',owned_excel_pid=pid,owned_excel_creation_time=psutil.Process(pid).create_time(),prior_pids=sorted(prior),modules={},dll_sha256={p.name:sha(p) for p in DELIVERY.glob('*.dll')})
    def save():
        (DELIVERY/'WORKBOOK_VALIDATION.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    save()
    book=None;component=None;control=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        names=[component.Name for component in book.VBProject.VBComponents if component.Type==1]
        for name in names: book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
        for source in sorted((ROOT/'src/vba').glob('*.bas')):
            component=book.VBProject.VBComponents.Add(1);component.Name=source.stem;component.CodeModule.AddFromString(text(source))
            report['modules'][source.name]={'source_sha256':sha(source),'loaded_code_sha256':hashlib.sha256(component.CodeModule.Lines(1,component.CodeModule.CountOfLines).encode('utf-8')).hexdigest()}
        assert len(report['modules'])==62
        compile_book(app,book);report['compile']='PASS';save()
        call=lambda name,*values:app.Run("'"+book.Name+"'!"+name,*values)
        call('RPX_Setup')
        assert call('RPX_Setting','ACCURACY_DLL_RECOVERY',-1)==0
        inputs={name:book.Worksheets(name).UsedRange.Value2 for name in ['要素定義','材料データ','浸透入力']}
        policy=call('RPX_Setting','ANALYSIS_POLICY','')
        component=book.VBProject.VBComponents.Add(1);component.Name='DistributionPresetProbe'
        component.CodeModule.AddFromString('''
Public Function DistributionPresetProbeRun() As String
    Dim oldLower As RPX_MeshState, oldUpper As RPX_MeshState
    oldLower = RBestLower: oldUpper = RBestUpper
    On Error GoTo Failed
    ReDim RBestLower.field(0 To 1): ReDim RBestUpper.field(0 To 1)
    RBestLower.field(0) = 0.123456789012345: RBestLower.field(1) = -7.125
    RBestUpper.field(0) = 0.987654321098765: RBestUpper.field(1) = 9.25
    RPX_AccuracyPreset
    If RBestLower.field(0) <> 0.123456789012345 Or RBestLower.field(1) <> -7.125 Then Err.Raise 5, , "PRESET_CHANGED_LOWER_FIELD"
    If RBestUpper.field(0) <> 0.987654321098765 Or RBestUpper.field(1) <> 9.25 Then Err.Raise 5, , "PRESET_CHANGED_UPPER_FIELD"
    DistributionPresetProbeRun = "PASS"
CleanUp:
    RBestLower = oldLower: RBestUpper = oldUpper
    Exit Function
Failed:
    DistributionPresetProbeRun = "FAIL " & Err.Description
    Resume CleanUp
End Function
''')
        preset_result=call('DistributionPresetProbeRun')
        assert preset_result=='PASS',preset_result
        report['preset_preserves_saved_field_vectors']=True
        book.VBProject.VBComponents.Remove(component);component=None
        compile_book(app,book)
        expected={'MODE':'Fs','TARGET':1000,'QUADRATURE':8,'FS_TOL':.0001,'GAP':.01,'ADAPT_ACCURACY_CAP':24000,'ADAPT_LOWER_TARGET':4000,'ADAPT_CERT_QUADRATURE':8,'ACCURACY_DLL_RECOVERY':1,'ACCURACY_DLL_CONE_SCALING':1,'HSD_SOLVE_SECONDS':300}
        report['precision_preset']={key:call('RPX_Setting',key,-1) for key in expected}
        assert report['precision_preset']==expected
        assert policy==call('RPX_Setting','ANALYSIS_POLICY','')
        assert inputs=={name:book.Worksheets(name).UsedRange.Value2 for name in inputs}
        report['preset_preserves_input_and_policy']=True
        save()
        call('RPX_Setup')
        book.Worksheets('浸透結果').Cells.Clear()
        guide=book.Worksheets('実行案内')
        guide.Cells.Clear()
        for row,message in enumerate(['RPFEM G1R18', 'P06の1%目標は継続課題です。', '64bit Excelと同じフォルダー内の2個のDLLを使用します。', '追加の精度回復は初期設定で無効です。', '設定シートの「精度重視の設定」で追加候補とq8再検証を有効にします。', '手順・制限・未達状況は同梱のREADME_JA.mdをご確認ください。'],1):
            guide.Cells(row,1).Value2=message
        guide.Columns('A').ColumnWidth=100
        guide=None
        book.Worksheets('解析結果').Range('A1').Value2='G1R18 未解析（P06の1%目標は継続課題）'
        book.BuiltinDocumentProperties('Title').Value='RPFEM G1R18'
        book.Save();book.Close(SaveChanges=False);book=None
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        compile_book(app,book);report['reopen_compile']='PASS'
        assert call('RPX_Setting','ACCURACY_DLL_RECOVERY',-1)==0
        assert call('RPX_Setting','ADAPT_CERT_QUADRATURE',-1)==0
        assert str(book.Worksheets('解析結果').Range('A1').Value2).startswith('G1R18 未解析')
        assert {c.Name for c in book.VBProject.VBComponents if c.Type==1}=={p.stem for p in (ROOT/'src/vba').glob('*.bas')}
        report['initial_state']='UNSOLVED;additional_recovery_OFF;source_modules_only'
        report['state']='PASS'
    except BaseException as exc:
        report['state']='FAIL';report['error']=repr(exc)
        raise
    finally:
        if book is not None:book.Close(SaveChanges=False)
        app.Quit();component=None;book=None;control=None;app=None;gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        report['owned_excel_exited']=not psutil.pid_exists(pid)
        report['workbook_sha256']=sha(path)
        save()
    return report


def package():
    files={}
    def add(relative,source):
        assert relative not in files
        files[relative]=source.read_bytes()
    for name in [NAME+'.xlsm','rpfem_native_x64.dll','rpfem_accuracy_x64.dll','WORKBOOK_VALIDATION.json']:
        add(name,DELIVERY/name)
    for name in ['README_JA.md','SOURCE_MANIFEST.json']:add(name,ROOT/name)
    for directory in ['src/vba','import','tools']:
        for path in sorted((ROOT/directory).glob('*')):
            if path.is_file():add(str(path.relative_to(ROOT)).replace('\\','/'),path)
    for path in sorted((BASE/'native').rglob('*')):
        if path.is_file() and path.suffix.lower() in ['.c','.cpp','.h','.asm','.cmd','.txt']:
            add('native/baseline/'+str(path.relative_to(BASE/'native')).replace('\\','/'),path)
    accuracy=ROOT/'native/accuracy'
    for path in sorted(accuracy.rglob('*')):
        relative=path.relative_to(accuracy)
        if not path.is_file() or relative.parts[0] in ['target','.gitignore','validation']:continue
        if relative.parts[0]=='bin' and path.name not in ['rpfem_accuracy_x64.dll','build_manifest.json']:continue
        add('native/accuracy/'+str(relative).replace('\\','/'),path)
    manifest={'release':'G1R18','p06_target_complete':False,'runtime_requires_python':False,'byte_scope':'Exact generated files inside this ZIP and release Git blobs (-text attributes)','files':{name:{'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()} for name,data in sorted(files.items())}}
    raw=(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n').encode('utf-8')
    (DELIVERY/'MANIFEST.json').write_bytes(raw);files['MANIFEST.json']=raw
    archive=DELIVERY/(NAME+'.zip')
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as output:
        for name,data in sorted(files.items()):
            item=zipfile.ZipInfo(NAME+'/'+name,(2026,10,14,0,0,0));item.compress_type=zipfile.ZIP_DEFLATED;output.writestr(item,data)
    with zipfile.ZipFile(archive) as result:
        assert len(result.namelist())==len(files)
        for name,data in files.items():assert result.read(NAME+'/'+name)==data
    report={'status':'PASS','archive':str(archive),'archive_sha256':sha(archive),'bytes':archive.stat().st_size,'entries':len(files),'all_member_bytes_verified':True}
    (DELIVERY/'ZIP_VALIDATION.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(report),flush=True)


if __name__=='__main__':
    print(json.dumps(create_book(),ensure_ascii=False),flush=True)
    package()
