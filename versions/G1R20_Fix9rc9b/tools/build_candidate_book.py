"""Compile/save/reopen an unpublished combined candidate; save verified speed preset."""
from pathlib import Path
import argparse, gc, hashlib, importlib.util, json, re, shutil, time, uuid
import psutil

ROOT=Path(__file__).resolve().parents[1]
BASE=ROOT.parent/'G1R20_Fix4'
spec=importlib.util.spec_from_file_location('candidate_sources',ROOT/'tools/current_sources.py')
sources=importlib.util.module_from_spec(spec);spec.loader.exec_module(sources)


def code(path):
    return '\r\n'.join(s for s in path.read_text(encoding='utf8').splitlines() if not s.startswith('Attribute VB_'))


def code_key(text):
    # Excel changes identifier capitalization across the project on serialization.
    # VBE also moves indentation past a leading label's colon. Normalize only
    # that formatter change; preserve literals/comments and remaining spacing.
    lines=[]
    for line in text.replace('\r\n','\n').rstrip('\n').split('\n'):
        line=re.sub(r'^[ \t]*([A-Za-z_]\w*):[ \t]*',r'\1:',line)
        result=[];i=0
        while i<len(line):
            if line[i]=="'":
                result.append(line[i:]);break
            if line[i]=='"':
                begin=i;i+=1
                while i<len(line):
                    if line[i]=='"':
                        if i+1<len(line) and line[i+1]=='"':i+=2;continue
                        i+=1;break
                    i+=1
                result.append(line[begin:i]);continue
            result.append(line[i].lower());i+=1
        lines.append(''.join(result))
    return hashlib.sha256('\n'.join(lines).encode('utf8')).hexdigest()


def main():
    import win32com.client, win32process
    parser=argparse.ArgumentParser();parser.add_argument('--output-root',type=Path,required=True);args=parser.parse_args()
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError('Another Excel exists: '+repr(prior))
    out=args.output_root.resolve();sources.generate(out)
    manifest=json.loads((out/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    delivery=out/'delivery';delivery.mkdir()
    path=delivery/'RPFEM_20261021_G1R20_Fix9rc9b.xlsm'
    shutil.copyfile(BASE/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm',path)
    shutil.copyfile(BASE/'delivery/rpfem_native_x64.dll',delivery/'rpfem_native_x64.dll')
    for name in ['README_JA.md','TEST_RESULTS.md']:
        if (ROOT/name).exists():shutil.copyfile(ROOT/name,out/name)
    rec=dict(state='PREPARING',run_token=uuid.uuid4().hex,prior_pids=prior,
             candidate_only=True,published_files_modified=False,solver_calls=0,
             harness_sha256=sources.sha(__file__),baseline_workbook_sha256=sources.sha(BASE/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm'),
             dll_sha256=sources.sha(delivery/'rpfem_native_x64.dll'),modules={},workbook_modules={},
             source_bundle_sha256=sources.bundle_key(manifest))
    save=lambda:(delivery/'WORKBOOK_VALIDATION.json').write_text(json.dumps(rec,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    save();shutil.copyfile(__file__,out/'build_candidate_book_snapshot.py')
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    rec.update(owned_excel_pid=pid,owned_excel_creation_time=psutil.Process(pid).create_time());save()
    book=component=control=c=None
    def compile_book():
        nonlocal control
        book.Activate();book.VBProject.VBComponents('RPX_AccuracyNative').CodeModule.CodePane.Show()
        control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.3)
        if control.Enabled:raise RuntimeError('Candidate compile failed')
        app.VBE.MainWindow.Visible=False
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        expected={p.stem for p in (out/'src/vba').glob('*.bas')}
        if len(expected)!=63:raise ValueError('Full source must contain 63 modules')
        for n in [c.Name for c in book.VBProject.VBComponents if c.Type==1]:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(n))
        for p in sorted((out/'src/vba').glob('*.bas')):
            component=book.VBProject.VBComponents.Add(1);component.Name=p.stem;component.CodeModule.AddFromString(code(p))
            loaded=component.CodeModule.Lines(1,component.CodeModule.CountOfLines)
            expected_key=code_key(code(p))
            rec['modules'][p.name]=dict(source_sha256=sources.sha(p),source_code_key_sha256=expected_key,
                loaded_code_sha256=hashlib.sha256(loaded.encode('utf8')).hexdigest(),loaded_code_key_sha256=code_key(loaded))
            if code_key(loaded)!=expected_key:raise ValueError('Loaded code differs from its declared source: '+p.name)
        for name,row in manifest['workbook_modules'].items():
            p=out/row['utf8_path'];component=book.VBProject.VBComponents(name)
            if component.Type!=100:raise ValueError('Expected workbook document module: '+name)
            if component.CodeModule.CountOfLines:component.CodeModule.DeleteLines(1,component.CodeModule.CountOfLines)
            component.CodeModule.AddFromString(code(p))
            loaded=component.CodeModule.Lines(1,component.CodeModule.CountOfLines);expected_key=code_key(code(p))
            rec['workbook_modules'][name]=dict(source_path=row['utf8_path'],source_sha256=sources.sha(p),
                source_code_key_sha256=expected_key,loaded_code_sha256=hashlib.sha256(loaded.encode('utf8')).hexdigest(),
                loaded_code_key_sha256=code_key(loaded))
            if code_key(loaded)!=expected_key:raise ValueError('Loaded workbook code differs: '+name)
        compile_book();rec['compile']='PASS';save()
        call=lambda name,*v:app.Run("'"+book.Name+"'!"+name,*v)
        physical_before={n:book.Worksheets(n).UsedRange.Value2 for n in ['要素定義','材料データ','浸透入力']}
        call('RPX_Setup')
        if physical_before!={n:book.Worksheets(n).UsedRange.Value2 for n in physical_before}:raise ValueError('Setup changed physical workbook input')
        ordinary={'DIRECT_DLL':0,'ADAPT_PRECISION_FROM_START':0,'ADAPT_C0_WEDGE_AUTO':0,'GAP':.02}
        rec['ordinary_defaults']={k:call('RPX_Setting',k,-1) for k in ordinary}
        if rec['ordinary_defaults']!=ordinary:raise ValueError('Ordinary defaults changed')
        inputs={n:book.Worksheets(n).UsedRange.Value2 for n in ['要素定義','材料データ','浸透入力']}
        policy=call('RPX_Setting','ANALYSIS_POLICY','')
        call('RPX_Speed2pctPreset')
        speed={'GAP':.02,'QUADRATURE':8,'FS_TOL':.0001,'DIRECT_DLL':1,'ADAPT_C0_WEDGE_AUTO':1,'ADAPT_PRECISION_FROM_START':1,'ACCURACY_DLL_RECOVERY':1,'LOWER_MODEL_CACHE':0,'TARGET':1000}
        rec['speed_preset']={k:call('RPX_Setting',k,-1) for k in speed}
        if rec['speed_preset']!=speed:raise ValueError('Preset differs from the audited cold profile')
        if inputs!={n:book.Worksheets(n).UsedRange.Value2 for n in inputs} or policy!=call('RPX_Setting','ANALYSIS_POLICY',''):raise ValueError('Preset changed physical inputs or policy')
        rec['preset_preserves_physical_input_and_policy']=True
        call('RPX_G20ImprovementPreset')
        improvement={'ADAPT_BALANCED_LOWER_TARGET':1000,'ADAPT_UPPER_FLOW_PASSES':5,'FS_TERMINAL_SCHEDULE':1,'ACCURACY_DLL_UPPER_RECENTER':1,'ACCURACY_DLL_DUAL_LINEAR_REPAIR':1,'ACCURACY_DLL_UPPER_LP_SECONDS':30,'DIRECT_DLL_SECONDS':30,'ACCURACY_DLL_DUAL_SECONDS':60,'ACCURACY_DLL_DUAL_REGULARIZATION':1e-8,'ACCURACY_DLL_DUAL_EQUILIBRATE':1,'ACCURACY_DLL_EQUILIBRATE':0,'ACCURACY_DLL_SECONDS':90}
        rec['opt_in_improvement_profile']={k:call('RPX_Setting',k,-1) for k in improvement}
        if rec['opt_in_improvement_profile']!=improvement:raise ValueError('Improvement profile differs')
        if inputs!={n:book.Worksheets(n).UsedRange.Value2 for n in inputs} or policy!=call('RPX_Setting','ANALYSIS_POLICY',''):raise ValueError('Improvement profile changed physical input')
        call('RPX_Speed2pctPreset')
        flags={'ADAPT_BALANCED_LOWER_TARGET':0,'ADAPT_UPPER_FLOW_PASSES':0,'FS_TERMINAL_SCHEDULE':0,'ACCURACY_DLL_UPPER_RECENTER':0,'ACCURACY_DLL_DUAL_LINEAR_REPAIR':0,'ACCURACY_DLL_DUAL_REGULARIZATION':0,'DIRECT_DLL_SECONDS':5,'ACCURACY_DLL_DUAL_SECONDS':20,'ACCURACY_DLL_SECONDS':120}
        rec['saved_optional_flags']={k:call('RPX_Setting',k,-1) for k in flags}
        if rec['saved_optional_flags']!=flags:raise ValueError('Optional profile failed to reset')
        # Fix5: accuracy first. Save with the measured 2% standard preset (Fix4 26-case
        # results); the G1R20 improvement candidate stays opt-in until verified.
        call('RPX_Speed2pctPreset')
        rec['final_saved_speed_profile']={k:call('RPX_Setting',k,-1) for k in speed}
        if rec['final_saved_speed_profile']!=speed:raise ValueError('Final saved 2% profile differs')

        # The ordinary API defaults remain opt-in; this candidate is saved with
        # the speed preset explicitly selected, matching the verified cold runs.
        rec['saved_profile']='Speed2pctPreset (2% standard); GAP.02/q8/direct1/precision1/auto1/accuracy_recovery1'
        rec['runtime_available']=bool(call('RPX_NativeAvailable'))
        if not rec['runtime_available']:raise ValueError('Unified runtime unavailable')
        rec['native_memory_maps']=[m.path for m in psutil.Process(pid).memory_maps() if Path(m.path).name.lower().startswith('rpfem_')]
        if rec['native_memory_maps']!=[str(delivery/'rpfem_native_x64.dll')]:raise ValueError('Unexpected DLL mapping')
        guide=book.Worksheets('実行案内');guide.Cells(1,1).Value2='G1R20 Fix9 rc9b 候補（未公開・未検証）'
        guide.Cells(4,1).Value2='GAPの標準・合格基準は2%です。2%標準の設定（q8・直接DLL・精度回復）を選択済み。Fix9 rc9b候補（未公開・未検証）: 上界の元問題のDLL解法で平衡化（ACCURACY_DLL_UPPER_EQUILIBRATE=1、0でrc9aと同じ）、双対診断ログの修正（判定・許容差は変更なし）。以下はrc9aの内容: 不具合修正（大きい上界の段で解が得られないまま時間切れになったときのエラー9を修正し、後段でエラーが起きても確定済みの下界・上界とFsを書き出す）。以下はrc8の内容: 薄層チェーンのDLL解法で平衡化（equilibrate=1、ADAPT_THIN_DLL_EQUILIBRATE）、下界の追加段を既定3段に、時間保護をDLL時間1.15倍と追加段ごとの予約に対応。以下はrc7bの内容: 薄層チェーンのみ、同じメッシュ上のFs試行間でHSDの状態を引き継ぐウォームスタート（内部へ再中心化、停滞時は同じ予算内でコールドに戻す）。以下はrc7の内容: 薄層チェーンのみ、大きい上界のHSDを1800秒・600反復、下界HSD 1200秒、根探索の初期区間を前段の値の近くに狭めて評価回数を削減、上界の分割を要素当たり効果順、狙いを絞った下界追加1段。以下はrc6cの内容: 薄層チェーンの既定時計を18171秒に（ADAPT_THIN_NOMINAL_SECONDS）。以下はrc6の内容: 薄層チェーンのみ、上界の1段の要素数を3000・増加1.6倍までに制限、大きい上界でDLL状態5なら1回再試行・HSDの時間と反復を拡大、残り時間をケース制限（ADAPT_THIN_CASE_LIMIT_SECONDS）から計算、残り時間で下界の追加段（最大3）。以下はrc5の内容: 上界の根探索でDLL状態5を「根より下」の符号として扱いHSDに戻らない、失敗した段の有効な最良解を保持、HSD制限時間を残り時間に合わせる、大きい下界メッシュはθを緩める、各側1回の追加再試行。以下はrc4bの内容: 早期開始したときだけ上界チェーンをθ0.95・要素数上限7000にします（ADAPT_THIN_EARLY_UPPER_THETA、ADAPT_THIN_EARLY_UPPER_CAP）。以下はrc3bの内容: 早期開始は条件付き（ADAPT_THIN_EARLY_START=1：上界10段以上・下界3段以上・上限/粗い要素数2倍以上などチェーンの余裕が十分なときだけ、2=常に、0=無効）。補助の上界解はAlmostSolvedも指標に使用。下界チェーンの要素数上限6000。以下はFix9の内容: Fix8の薄層チェーンを試作に合わせました（中線を独立した最初の段階に、下界の指標を下界メッシュ上の上界解から、入替え前に上界を解き直す、下界の根の失敗で1回再試行、チェーン内のHSD・DLLの制限時間を規模に合わせる）。以下はFix8の内容: 薄い材料層（厚さ比<0.1）を自動検出した場合のみ、上界・下界を別メッシュで改良（層の中線分割、方向付き分割＋辺フリップ、混合指標＋扇形分割）。下界の等式残差の直接修復（荷重固定）。設定 ADAPT_THIN_LAYER=0 で無効。Fix7の内容を含む。P06N/P06Aの検証は別PCで実施予定です。'
        guide=None
        book.Worksheets('解析結果').Range('A1').Value2='G1R20 Fix9 rc9b 未解析'
        book.Save();book.Close(False);book=None
        book=app.Workbooks.Open(str(path),UpdateLinks=0);compile_book()
        if {c.Name for c in book.VBProject.VBComponents if c.Type==1}!=expected:raise ValueError('Unexpected reopened source modules')
        rec['reopened_speed_profile']={k:call('RPX_Setting',k,-1) for k in speed}
        if rec['reopened_speed_profile']!=speed:raise ValueError('Saved speed profile changed on reopen')
        if physical_before!={n:book.Worksheets(n).UsedRange.Value2 for n in physical_before}:raise ValueError('Reopened physical input changed')
        if policy!=call('RPX_Setting','ANALYSIS_POLICY',''):raise ValueError('Reopened analysis policy changed')
        rec['reopened_physical_input_policy_unchanged']=True
        for p in (out/'src/vba').glob('*.bas'):
            c=book.VBProject.VBComponents(p.stem)
            reopened=c.CodeModule.Lines(1,c.CodeModule.CountOfLines)
            rec['modules'][p.name]['reopened_code_sha256']=hashlib.sha256(reopened.encode('utf8')).hexdigest()
            rec['modules'][p.name]['reopened_code_key_sha256']=code_key(reopened)
            if code_key(reopened)!=rec['modules'][p.name]['loaded_code_key_sha256']:raise ValueError('Reopened code differs beyond identifier case: '+p.name)
        for name,row in manifest['workbook_modules'].items():
            c=book.VBProject.VBComponents(name);reopened=c.CodeModule.Lines(1,c.CodeModule.CountOfLines)
            rec['workbook_modules'][name]['reopened_code_sha256']=hashlib.sha256(reopened.encode('utf8')).hexdigest()
            rec['workbook_modules'][name]['reopened_code_key_sha256']=code_key(reopened)
            if code_key(reopened)!=rec['workbook_modules'][name]['loaded_code_key_sha256']:raise ValueError('Reopened workbook code differs: '+name)
        rec.update(state='PASS',reopen_compile='PASS',standard_modules=63,workbook_module_count=1,runtime_dll_count=1,
                   reopened_code_comparison='Identifier case and leading-label indentation normalized; string literals/comments/numbers and other characters exact; raw pre/post SHA retained')
    except BaseException as e:rec.update(state='FAIL',error=repr(e));save();raise
    finally:
        if book is not None:book.Close(False)
        app.Quit();book=component=control=c=app=None;gc.collect()
        for _ in range(200):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        rec['owned_excel_exited']=not psutil.pid_exists(pid)
        rec['workbook_sha256']=sources.sha(path);save()
    print(json.dumps({'state':rec['state'],'output':str(out),'workbook_sha256':rec['workbook_sha256'],'owned_excel_exited':rec['owned_excel_exited']}))

if __name__=='__main__':main()
