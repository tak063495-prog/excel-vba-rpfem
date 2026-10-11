from pathlib import Path
import json, hashlib, difflib
import win32com.client, win32process

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT.parent / 'G1R15/examples/RPFEM_20261011_G1R15_Levee_Example.xlsm'
ROOT.joinpath('results').mkdir(parents=True, exist_ok=True)
app = win32com.client.DispatchEx('Excel.Application')
book = None
try:
    app.Visible = False
    book = app.Workbooks.Open(str(SOURCE), UpdateLinks=0, ReadOnly=True)
    data = {'excel_version': app.Version, 'macro_security_unmodified': app.AutomationSecurity,
            'source_sha256': hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
            'pid': win32process.GetWindowThreadProcessId(app.Hwnd)[1], 'sheets': [], 'source_matches': {}}
    for sheet in book.Worksheets:
        data['sheets'].append(sheet.Name)
        if sheet.Name == '設定':
            data['settings'] = sheet.Range('A1:F20').Value2
        if sheet.Name == '浸透入力':
            data['seepage'] = sheet.Range('A1:Q10').Value2
        if sheet.Name == '解析結果':
            data['saved_results'] = sheet.Range('A1:H20').Value2
    extracted = ROOT / 'baseline_sources'
    extracted.mkdir(exist_ok=True)
    differences = []
    for component in book.VBProject.VBComponents:
        if component.Type != 1: continue
        text = component.CodeModule.Lines(1, component.CodeModule.CountOfLines) if component.CodeModule.CountOfLines else ''
        (extracted / (component.Name + '.bas')).write_text('Attribute VB_Name = "' + component.Name + '"\n' + text.replace('\r\n','\n'), encoding='utf-8')
        file = ROOT.parent / 'G1R15/src/vba' / (component.Name + '.bas')
        if file.exists():
            expected = '\n'.join(line for line in file.read_text(encoding='utf-8-sig').splitlines() if not line.startswith('Attribute VB_')).strip()
            data['source_matches'][component.Name] = text.replace('\r\n', '\n').strip() == expected
            if not data['source_matches'][component.Name]:
                differences += list(difflib.unified_diff(expected.splitlines(),text.replace('\r\n','\n').strip().splitlines(),fromfile='repo/'+component.Name,tofile='workbook/'+component.Name))
    ROOT.joinpath('results/source_discrepancies.diff').write_text('\n'.join(differences), encoding='utf-8')
    ROOT.joinpath('results/probe.json').write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    print(json.dumps({'version': data['excel_version'], 'security': data['macro_security_unmodified'],
                      'mismatches': [k for k,v in data['source_matches'].items() if not v]},ensure_ascii=False))
finally:
    if book is not None: book.Close(False)
    app.Quit()
