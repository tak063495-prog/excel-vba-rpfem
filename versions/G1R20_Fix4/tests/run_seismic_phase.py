"""Owned Excel history regression: actual Fs solves, then public mesh/run paths.

Default reproduces on the immutable Fix3 workbook. --candidate imports all Fix4
product modules into a disposable copy. --final-book tests the frozen saved Fix4
book without replacing product modules. No field, root or solver status is set.
"""
from pathlib import Path
import argparse, gc, hashlib, json, math, shutil, threading, time
import psutil, win32api, win32com.client, win32con, win32gui, win32process

HERE=Path(__file__).resolve().parent
ROOT=HERE.parent
REPO=ROOT.parents[1]
BASE=REPO/'versions/G1R20_Fix3'
def sha(p): return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def code(p): return '\r\n'.join(x for x in Path(p).read_text(encoding='utf8').splitlines() if not x.startswith('Attribute VB_'))
def main():
    parser=argparse.ArgumentParser()
    choice=parser.add_mutually_exclusive_group()
    choice.add_argument('--candidate',action='store_true'); choice.add_argument('--final-book',action='store_true')
    parser.add_argument('--expected-book-sha'); parser.add_argument('--production',action='store_true')
    parser.add_argument('--input-matrix',action='store_true'); parser.add_argument('--legacy',action='store_true'); args=parser.parse_args()
    if args.final_book and not args.expected_book_sha: parser.error('--final-book requires --expected-book-sha')
    fixed=args.candidate or args.final_book
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior: raise RuntimeError(f'Excel exists {prior}')
    out=ROOT/'results'/('seismic_candidate_' if fixed else 'seismic_negative_')/str(time.time_ns()); out.mkdir(parents=True)
    source_book=(ROOT/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm' if args.final_book else BASE/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm')
    if args.expected_book_sha: assert sha(source_book)==args.expected_book_sha, 'Selected frozen book changed'
    bookpath=out/'seismic.xlsm'; shutil.copyfile(source_book,bookpath)
    source_dll=(ROOT if args.final_book else BASE)/'delivery/rpfem_native_x64.dll'
    shutil.copyfile(source_dll,out/'rpfem_native_x64.dll')
    for p in [Path(__file__),HERE/'SeismicPhaseFix4Probe.bas',BASE/'tests/read_model_shear_c10.txt']: shutil.copyfile(p,out/p.name)
    if args.input_matrix or args.legacy: shutil.copyfile(HERE/'InputMatrixFix4Probe.bas',out/'InputMatrixFix4Probe.bas')
    product=ROOT if fixed else BASE
    observed=list((product/'src/vba').glob('*.bas'))+[product/'src/workbook/ThisWorkbook.code.txt']
    before={str(p):sha(p) for p in observed}
    rec=dict(status='RUNNING',candidate=fixed,final_book=args.final_book,product_modules_replaced=args.candidate,source_book=str(source_book),source_book_sha256=sha(source_book),dll_sha256=sha(source_dll),book_sha=sha(BASE/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm'),prior_excel=prior,calls={},dialogs=[],source_sha256=before)
    def save(): (out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2,allow_nan=False)+'\n',encoding='utf8')
    save(); app=win32com.client.DispatchEx('Excel.Application'); pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]; born=psutil.Process(pid).create_time()
    command=psutil.Process(pid).cmdline()
    process_handle=win32api.OpenProcess(0x1000|0x100000,False,pid)
    rec.update(owned_excel_pid=pid,owned_excel_birth=born,owned_excel_command=command); stop=threading.Event()
    def watchdog():
        if stop.wait(240): return
        rec['watchdog_timeout']=True; save()
        if psutil.pid_exists(pid):
            proc=psutil.Process(pid)
            if proc.create_time()==born and proc.name().upper()=='EXCEL.EXE' and proc.cmdline()==command: proc.terminate()
    monitor=threading.Thread(target=watchdog,daemon=True); monitor.start(); book=None; component=module=control=None
    def dialogs():
        while not stop.wait(.2):
            def visit(hwnd,_):
                if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd)!='#32770': return
                children=[]
                win32gui.EnumChildWindows(hwnd,lambda child,_:children.append((child,win32gui.GetWindowText(child))),None)
                words=[value for _,value in children]
                if words not in rec['dialogs']: rec['dialogs'].append(words); save()
                for child,label in children:
                    if label in ['OK','&OK'] or ('Visual Basic' in win32gui.GetWindowText(hwnd) and ('終了' in label or '&End' in label)):
                        win32gui.PostMessage(child,win32con.BM_CLICK,0,0)
            try: win32gui.EnumWindows(visit,None)
            except Exception as e:
                key='transient_window_1400' if getattr(e,'winerror',e.args[0] if e.args else None)==1400 else 'window_monitor_errors'
                if repr(e) not in rec.setdefault(key,[]): rec[key].append(repr(e))
    threading.Thread(target=dialogs,daemon=True).start()
    try:
        app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False; app.MultiThreadedCalculation.Enabled=False
        rec['workbooks_before_open']=[app.Workbooks(i).Name for i in range(1,app.Workbooks.Count+1)]
        assert not rec['workbooks_before_open'], 'Unexpected recovered/foreign workbook in the owned instance'
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        rec['opened_workbook_fullname']=book.FullName
        assert Path(book.FullName).resolve()==bookpath.resolve()
        if args.candidate:
            root=ROOT
            rec['source_modules']={}
            for p in sorted((root/'src/vba').glob('*.bas')):
                module=book.VBProject.VBComponents(p.stem).CodeModule
                if module.CountOfLines: module.DeleteLines(1,module.CountOfLines)
                module.AddFromString(code(p)); rec['source_modules'][p.name]=sha(p)
            p=root/'src/workbook/ThisWorkbook.code.txt'; module=book.VBProject.VBComponents('ThisWorkbook').CodeModule
            if module.CountOfLines: module.DeleteLines(1,module.CountOfLines)
            module.AddFromString(code(p)); rec['event_source_sha']=sha(p)
        component=book.VBProject.VBComponents.Add(1); component.Name='SeismicPhaseFix4Probe'; component.CodeModule.AddFromString(code(HERE/'SeismicPhaseFix4Probe.bas'))
        if args.input_matrix or args.legacy:
            component=book.VBProject.VBComponents.Add(1); component.Name='InputMatrixFix4Probe'; component.CodeModule.AddFromString(code(HERE/'InputMatrixFix4Probe.bas'))
        book.Activate(); component.CodeModule.CodePane.Show(); control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled: control.Execute(); time.sleep(.25)
        if control.Enabled: raise RuntimeError('Compile failure')
        app.VBE.MainWindow.Visible=False; rec['compile']='PASS'; save()
        def call(label,fn,*values):
            answer=app.Run("'"+book.Name+"'!"+fn,*values)
            row=list(answer)
            if fn in ['PhasePublicMesh','PhasePublicRun']:
                row[11]=hashlib.sha256(row[11].encode('utf-16-le')).hexdigest()
            rec['calls'][label]=row; save(); print(json.dumps(dict(label=label,result=row),ensure_ascii=False),flush=True)
            return row
        def loads_equal(row, expected): return all(math.isclose(a,b,rel_tol=0.,abs_tol=1e-12) for a,b in zip(row[7:11],expected))
        if args.input_matrix:
            accepted=['valid','c0','gamma0','davis','negative_coordinates','numeric_text_id','last_point_row','last_material_row','adjacent_regions','contained_regions','contained_touching_region']
            rejected=['duplicate_point','zero_point','negative_point','fraction_point','boolean_point','text_point','error_point','missing_referenced_point','duplicate_material','negative_c','negative_gamma','phi89','psi_above_phi','error_c','missing_material_ref','missing_rigid_ref','duplicate_region','identical_regions','repeated_polygon_point','self_cross_polygon','nonzero_self_cross_polygon','duplicate_face','missing_face_region','missing_face_point','wrap_mismatch','unknown_restraint','rigid_on_load_face','error_load']
            partial=['partial_point','partial_material','partial_rigid','partial_region','partial_face','aligned_overlap','aligned_overlap_narrow','aligned_overlap_left','overflow_point','overflow_material','overflow_rigid','overflow_region','overflow_face','boolean_load','boolean_x','boolean_c_false','boolean_gamma_false','boolean_psi_false']
            for case in accepted+rejected+partial:
                row=call(case,'InputMatrixCase',case)
                expected='ACCEPTED' if case in accepted or (case in partial and not fixed) else 'REJECTED'
                if row[0]!=expected: rec.setdefault('matrix_unexpected',[]).append(dict(case=case,expected=expected,result=row))
            rec['input_matrix_counts']=dict(valid=len(accepted),invalid=len(rejected),partial_id_missing=len(partial))
            rec['negative_partial_rows_accepted']=not fixed
            assert not rec.get('matrix_unexpected'), rec.get('matrix_unexpected')
        if args.legacy:
            valid=(out/'read_model_shear_c10.txt').read_text(encoding='utf8')
            files={
                'valid':valid,
                'fraction_node':valid.replace('0,1,2,0,-1,0,0,0,0','0.4,1.4,2.4,0,-1,0,0,0,0',1),
                'fraction_counts':valid.replace('4,2,1,0,4,1,1,8','4.4,2.4,1.4,0,4,1,1,8',1),
                'negative_c':valid.replace('10,0,0','-10,0,0',1),
                'unknown_boundary':valid.replace('"fixed"','"UNKNOWN"',1),
                'interior_boundary':valid.replace('0,1,"fixed"','0,2,"fixed"',1),
                'eof':'4,2,1,0,4,1,1,8\n10,0,0\n',
            }
            if fixed:
                first_triangle='0,1,2,0,-1,0,0,0,0'
                files.update({
                    'stress_scale_zero':valid.replace('4,2,1,0,4,1,1,8','4,2,1,0,4,0,1,8',1),
                    'strength_factor_zero':valid.replace('4,2,1,0,4,1,1,8','4,2,1,0,4,1,0,8',1),
                    'quadrature_zero':valid.replace('4,2,1,0,4,1,1,8','4,2,1,0,4,1,1,0',1),
                    'quadrature_fraction':valid.replace('4,2,1,0,4,1,1,8','4,2,1,0,4,1,1,8.4',1),
                    'negative_phi':valid.replace('10,0,0','10,-1,0',1),
                    'phi89':valid.replace('10,0,0','10,89,0',1),
                    'negative_gamma':valid.replace('10,0,0','10,0,-1',1),
                    'boolean_gamma':valid.replace('10,0,0','10,0,#FALSE#',1),
                    'fraction_material':valid.replace(first_triangle,'0,1,2,0.4,-1,0,0,0,0',1),
                    'missing_material':valid.replace(first_triangle,'0,1,2,1,-1,0,0,0,0',1),
                    'fraction_rigid':valid.replace(first_triangle,'0,1,2,0,-1.4,0,0,0,0',1),
                    'missing_rigid':valid.replace(first_triangle,'0,1,2,0,0,0,0,0,0',1),
                    'boolean_body':valid.replace(first_triangle,'0,1,2,0,-1,#TRUE#,0,0,0',1),
                    'fraction_boundary':valid.replace('0,1,"fixed"','0.4,1,"fixed"',1),
                    'missing_boundary_node':valid.replace('0,1,"fixed"','0,99,"fixed"',1),
                    'missing_boundary_rigid':valid.replace('"fixed",-1','"rigid",0',1),
                    'duplicate_boundary':valid.replace('1,2,"load",-1,0,0,0,1','0,1,"fixed",-1,0,0,0,0',1),
                    'overflow_number':valid.replace('10,0,0','1e309,0,0',1),
                })
                rows=valid.splitlines(); rigid='0.5,0.5,1,1,1,0,0,0,0,0,0'
                rigid_rows=[rows[0].replace('4,2,1,0,4,','4,2,1,1,4,')]+rows[1:8]+[rigid]+rows[8:]
                files['valid_rigid']='\n'.join(rigid_rows)+'\n'
                files['valid_nonrigid_sentinel0']=valid.replace('"fixed",-1','"fixed",0').replace('"load",-1','"load",0')
                files['rigid_fixed2']=files['valid_rigid'].replace(rigid,'0.5,0.5,2,1,1,0,0,0,0,0,0')
                files['rigid_fixed_fraction']=files['valid_rigid'].replace(rigid,'0.5,0.5,1.4,1,1,0,0,0,0,0,0')
                files['rigid_fixed_boolean']=files['valid_rigid'].replace(rigid,'0.5,0.5,#FALSE#,1,1,0,0,0,0,0,0')
                files['rigid_load_boolean']=files['valid_rigid'].replace(rigid,'0.5,0.5,1,1,1,#TRUE#,0,0,0,0,0')
            for name,value in files.items():
                path=out/('legacy_'+name+'.txt'); path.write_text(value,encoding='ascii',newline='\r\n')
                row=call('legacy_'+name,'LegacyInputCase',str(path))
                expected='ACCEPTED' if name.startswith('valid') or (name!='eof' and not fixed) else 'REJECTED'
                assert row[0]==expected,(name,row)
                assert row[13:15]==[not fixed,not fixed], 'Legacy read retained the prior current state'
                if name=='eof' and not fixed: assert row[3]!=0,'Expected leaked file handle'
                else: assert row[3]==0,'Reader retained a file handle'
            rec['legacy_cases']=len(files)
            busy=call('legacy_busy_gate','LegacyBusyCase',str(out/'legacy_valid.txt'))
            assert busy[0]==(5 if fixed else 0) and busy[2:]==[True,True], 'Busy legacy read changed the running state'
            if fixed: assert busy[1]=='ANALYSIS_IS_RUNNING'
        for q in [] if args.input_matrix or args.legacy else [2,8]:
            prefix=f'q{q}_'
            raw=call(prefix+'fresh_fixed','PhasePublicMesh','固定',.2,False,q)
            fs=call(prefix+'real_fs','PhaseRealFs',str(out/'read_model_shear_c10.txt'),str(out/f'actual_fs_q{q}.csv'))
            after=call(prefix+'after_fs_fixed','PhasePublicMesh','固定',.2,False,q)
            again=call(prefix+'repeated_fixed','PhasePublicMesh','固定',.2,False,q)
            assert raw[11]==after[11]==again[11], 'Physical worksheet inputs differ'
            assert fs[1]=='reference_total' and not fs[5], 'Actual Fs predecessor not established'
            assert raw[14]==after[14]==again[14]==q
            assert loads_equal(raw,[4.,-20.,0.,0.])
            if fixed:
                assert after[7:11]==raw[7:11]==again[7:11]
            else:
                assert loads_equal(after,[0.,-20.,4.,0.]) and again[7:11]==raw[7:11], 'Expected stale phase effect not reproduced'
                rec['negative_reproduced']=True
            call(prefix+'real_fs_before_reference','PhaseRealFs',str(out/'read_model_shear_c10.txt'),str(out/f'actual_fs_reference_q{q}.csv'))
            reference=call(prefix+'after_fs_reference','PhasePublicMesh','比例',.2,False,q)
            assert loads_equal(reference,[0.,0.,4.,-20.])
            zero=call(prefix+'zero_kh','PhasePublicMesh','固定',0.,False,q); assert loads_equal(zero,[0.,-20.,0.,0.])
        if args.production:
            fresh=call('public_fresh_bearing','PhasePublicRun','極限支持力',2)
            mainfs=call('public_main_fs','PhasePublicRun','Fs',2)
            bearing=call('public_after_fs_bearing','PhasePublicRun','極限支持力',2)
            assert mainfs[6]=='reference_total', 'Main Fs path did not leave the actual reference phase'
            assert not fresh[3] and not mainfs[3] and not bearing[3], 'Main solver did not finish'
            assert fresh[11]==bearing[11], 'Public run input drift'
            assert loads_equal(fresh,[4.,-20.,0.,0.])
            if fixed:
                assert loads_equal(bearing,[4.,-20.,0.,0.])
                assert all(math.isclose(a,b,rel_tol=1e-7,abs_tol=1e-7) for a,b in zip(fresh[12:14],bearing[12:14]))
            else: assert loads_equal(bearing,[0.,-20.,4.,0.])
        assert not rec['dialogs'] and not rec.get('window_monitor_errors')
        assert all(sha(p)==value for p,value in before.items()), 'Product source changed during execution'
        rec['status']='PASS'
    except BaseException as e:
        rec.update(status='FAIL',error=repr(e),owned_excel_alive_at_failure=psutil.pid_exists(pid),owned_excel_exit_code_at_failure=win32process.GetExitCodeProcess(process_handle)); save(); raise
    finally:
        component=module=control=None
        try:
            if book is not None: book.Close(False)
            app.Quit()
        except Exception as e: rec['cleanup_error']=repr(e)
        book=app=None; gc.collect(); stop.set()
        for _ in range(100):
            if not psutil.pid_exists(pid): break
            time.sleep(.1)
        rec['owned_excel_exited']=not psutil.pid_exists(pid)
        rec['owned_excel_exit_code']=win32process.GetExitCodeProcess(process_handle); process_handle.Close()
        rec['owned_excel_children']=[p.info for p in psutil.process_iter(['pid','ppid','name','create_time','cmdline']) if p.info['ppid']==pid and (p.info['name'] or '').upper()=='EXCEL.EXE' and p.info['create_time']>=born]
        rec['baseline_unchanged']=rec['book_sha']==sha(BASE/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm'); rec['source_book_unchanged']=rec['source_book_sha256']==sha(source_book); rec['source_files_unchanged']=all(sha(p)==value for p,value in before.items()); save()
    assert not rec['owned_excel_children'], 'Owned Excel spawned a recovery child'
    print(json.dumps(dict(status=rec['status'],out=str(out),owned_excel_exited=rec['owned_excel_exited'])),flush=True)
if __name__=='__main__': main()

