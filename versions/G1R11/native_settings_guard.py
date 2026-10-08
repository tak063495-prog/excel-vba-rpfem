from native_common import *

release = json.loads((R / 'Release_Verification.json').read_text(encoding='utf-8'))
production = Path(release['workbook'])
assert hashlib.sha256(production.read_bytes()).hexdigest() == release['sha256']
work = R / 'tests' / f'SettingsGuard_{time.time_ns()}.xlsm'
shutil.copy2(production, work)
app = win32com.client.DispatchEx('Excel.Application')
app.Visible = False
app.DisplayAlerts = False
book = None
probe = '''
Option Explicit
Public Function CheckSettingGuard(ByVal rowNumber As Long, ByVal newValue As Variant) As String
    On Error GoTo Failed
    XCancel = False
    Call RPX_InputBegin
    ThisWorkbook.Worksheets("設定").Cells(rowNumber, 2).Value2 = newValue
    RPX_AssertInputs "g1r11_setting_changed"
    CheckSettingGuard = "UNEXPECTED_SUCCESS"
    Call RPX_InputEnd
    Exit Function
Failed:
    CheckSettingGuard = "ERROR " & Err.Number & ":" & Err.Description
    Call RPX_InputEnd
    XCancel = False
End Function
'''
rows = []
try:
    book = app.Workbooks.Open(str(work), UpdateLinks=0)
    app.EnableEvents = False
    component = book.VBProject.VBComponents.Add(1)
    component.Name = 'SettingsGuardProbe'
    component.CodeModule.AddFromString(probe)
    compile_book(app, book)
    for row, value, key in [(48, 'LEGACY', 'ADAPT_MESH_POLICY'),
                            (49, 1001, 'ADAPT_REFINE_CAP'),
                            (50, 'LEGACY', 'FS_BRACKET_POLICY')]:
        ws = book.Worksheets('設定')
        previous = ws.Cells(row, 2).Value2
        result = app.Run(f"'{work.name}'!CheckSettingGuard", row, value)
        assert result.startswith('ERROR 5:') and 'INPUT_MISMATCH' in result, result
        assert ws.Cells(row, 2).Value2 == value
        rows.append({'test': key, 'result': result, 'events_disabled': True,
                     'user_changed_value_preserved': True,
                     'scope': 'NATIVE_INPUT_GUARD_ON_PRODUCTION_COPY'})
        ws.Cells(row, 2).Value2 = previous
    (R / 'results/native_settings_guard.json').write_text(
        json.dumps(rows, ensure_ascii=False, indent=2), encoding='utf-8')
    print(rows, flush=True)
finally:
    if book is not None:
        book.Close(False)
    app.Quit()
assert hashlib.sha256(production.read_bytes()).hexdigest() == release['sha256']

