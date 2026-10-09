from pathlib import Path
import difflib,json,hashlib
ROOT=Path(__file__).resolve().parents[1]
def edit(name,old,new):
    p=ROOT/'src/vba'/name;s=p.read_text(encoding='utf-8');assert s.count(old)==1,(name,old)
    p.write_text(s.replace(old,new),encoding='utf-8',newline='\n')
edit('RPX_Audit.bas','    CheckError Abs(RReferenceWork - 1#), 12, -1: RValue = RDissipation - RFixedWork',
'''    If RPX_ZeroCostUpperModel() Then
        RReferenceWork = RPX_PreciseUpperWork(xx)
        RPX_DiagEvent "upper_work_precision", "method=DD;W=" & RReferenceWork & ";tolerance=0.0000001"
    End If
    CheckError Abs(RReferenceWork - 1#), 12, -1: RValue = RDissipation - RFixedWork''')
edit('RPX_HSDCandidate.bas','    installed = True: xx = candidate.field: XObjective = candidate.objective',
'''    If candidate.mode = "upper" Then
        If RPX_NormalizeZeroCostUpper(candidate) Then candidate.field = candidate.x
    End If
    installed = True: xx = candidate.field: XObjective = candidate.objective''')
edit('RPX_Mesh.bas','    Call RPX_Topology: Call RPX_MapFaces: Call RPX_RepairFreeCorners\n    RPX_LoadPhase = "raw"',
'''    RPX_LoadPhase = "raw"
    Call RPX_Topology: Call RPX_MapFaces: Call RPX_RepairFreeCorners
    If RPX_RepairLoadSteps() Then Call RPX_AssignRegions''')
edit('RPX_Fs.bas','    FsNumericRecoverable = RPX_Recoverable(number, message, "fs_trial", False)',
'''    If Not XCancel And number = 5 And message = "HSD_DD_PRODUCT_UNDERFLOW" Then
        ' A failed point is UNKNOWN, never a sign/field/mesh proof. Retry only inside Fs search.
        XStatus = "NUMERICAL_UNKNOWN": RProgramKind = ""
        RPX_DiagEvent "fs_numerical_unknown", "reason=" & message & ";endpoints_unchanged=True;phase=fs_trial"
        FsNumericRecoverable = True
    Else
        FsNumericRecoverable = RPX_Recoverable(number, message, "fs_trial", False)
    End If''')
# Trial preparation is transactional: no partially normalized state can pass another gate.
p=ROOT/'src/vba/RPX_WorkPrecision.bas';s=p.read_text(encoding='utf-8')
start=s.index('Public Function RPX_NormalizeZeroCostUpper(')
head=s[:start];tail=s[start:]
body_start=tail.index('    w = RPX_PreciseUpperWork')
tail=tail[:body_start]+'    Dim trial As RPX_HsdCandidateState\n    trial = candidate\n'+tail[body_start:]
begin=tail.index('    w = RPX_PreciseUpperWork')
tail=tail[:begin]+tail[begin:].replace('candidate.','trial.')
tail=tail.replace('    RPX_NormalizeZeroCostUpper = True','    candidate = trial\n    RPX_NormalizeZeroCostUpper = True')
p.write_text(head+tail,encoding='utf-8',newline='\n')
# Source and strict lossless CP932/CRLF imports, plus unified diff.
changes=[];diff=[]
for p in sorted((ROOT/'src/vba').glob('*.bas')):
    b=ROOT/'before/vba'/p.name;old=b.read_text(encoding='utf-8') if b.exists() else ''
    new=p.read_text(encoding='utf-8');raw=('\r\n'.join(new.splitlines())+'\r\n').encode('cp932',errors='strict')
    assert raw.decode('cp932').splitlines()==new.splitlines()
    (ROOT/'import'/p.name).write_bytes(raw)
    if old!=new:
        changes.append(p.name);diff.extend(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='G1R12/'+p.name,tofile='G1R13/'+p.name))
(ROOT/'changes.diff').write_text(''.join(diff),encoding='utf-8')
(ROOT/'results/source_changes.json').write_text(json.dumps({'changed':changes,'source_modules':len(list((ROOT/'src/vba').glob('*.bas'))),'import_encoding':'CP932 strict lossless / CRLF'},indent=2),encoding='utf-8')
print(json.dumps(changes))
