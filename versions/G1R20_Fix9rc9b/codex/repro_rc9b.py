"""Operator-run rc9b P02A_THIN reproduction. Import/--help never opens Excel.

The source workbook and adjacent DLL are read only; each scenario uses private
copies below --out. No module injection, build, or macro-security changes.
The embedded case/configuration comes from codex_fix9rc9a_repro. E1 tests preset
defaults; E0 disables the new selector and retains rc9a's legacy thin behavior.
"""

import argparse
import csv
import gc
import hashlib
import json
import math
import multiprocessing as mp
import shutil
import time
import traceback
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime, timezone
from pathlib import Path

import psutil

ROOT = Path(__file__).resolve().parent
PRESET_MACRO = "RPX_Speed2pctPreset"
CASE_KEYS = ['GRAVITY', 'TEST_INJECTION', 'HSD_CAPTURE', 'DIAGNOSTIC_MODE',
             'ANALYSIS_POLICY', 'SEISMIC_KH', 'SEISMIC_DIR']
PRESET_EXPECTED = {
    'MODE': 'Fs', 'TARGET': 1000, 'QUADRATURE': 8, 'FS_TOL': 0.0001,
    'GAP': 0.02, 'ADAPT': 1, 'CYCLES': 8, 'RAW_LOWER_AUDIT': 1, 'C0_AUTO': 0,
    'ADAPT_ACCURACY_CAP': 24000, 'ADAPT_ACCURACY_ROUNDS': 8,
    'ADAPT_LOWER_TARGET': 4000, 'ADAPT_LOWER_FAN_ROUNDS': 1,
    'ADAPT_CERT_QUADRATURE': 8, 'ACCURACY_DLL_RECOVERY': 1,
    'ACCURACY_DLL_SECONDS': 120, 'ACCURACY_DLL_DUAL_SECONDS': 20,
    'ACCURACY_DLL_CONE_SCALING': 1, 'HSD_SOLVE_SECONDS': 300, 'DIRECT_DLL': 1,
    'ADAPT_PRECISION_FROM_START': 1, 'ADAPT_C0_WEDGE_AUTO': 1,
    'ADAPT_BALANCED_LOWER_TARGET': 0, 'ADAPT_UPPER_FLOW_PASSES': 0,
    'FS_TERMINAL_SCHEDULE': 0, 'ACCURACY_DLL_UPPER_RECENTER': 0,
    'ACCURACY_DLL_DUAL_LINEAR_REPAIR': 0, 'ACCURACY_DLL_DUAL_REGULARIZATION': 0,
    'DIRECT_DLL_SECONDS': 5, 'ACCURACY_DLL_UPPER_EQUILIBRATE': 1,
}
# Exactly the rc9a short numerical setup, with test injection disabled.
# Solver time budgets, tolerances, quadrature, and certificate gates are untouched.
SHORT = {
    'MODE': 'Fs', 'TARGET': 320, 'ADAPT': 1, 'CYCLES': 1,
    'TEST_INJECTION': 0, 'ACCURACY_DLL_RECOVERY': 1,
    'ADAPT_PRECISION_FROM_START': 1, 'ADAPT_THIN_LAYER': 2,
    'ADAPT_THIN_UPPER_ROUNDS': 1, 'ADAPT_THIN_UPPER_CAP': 800,
    'ADAPT_THIN_UPPER_ROUND_MAX_ELEMENTS': 800,
    'ADAPT_THIN_EARLY_UPPER_CAP': 800, 'ADAPT_THIN_LOWER_ROUNDS': 0,
    'ADAPT_THIN_LOWER_EXTRA_ROUNDS': 0, 'ADAPT_THIN_MID_SEPARATE': 0,
    'GAP': 0.000001, 'ADAPT_ACCURACY_CAP': 800,
    'ADAPT_ACCURACY_ROUNDS': 1, 'ADAPT_LOWER_TARGET': 0,
    'ADAPT_LOWER_FAN_ROUNDS': 0,
}
EVENTS = {
    'accuracy_dll_call', 'accuracy_dual_vector', 'accuracy_dll_original_dual',
    'direct_dll_begin', 'direct_dll_end', 'thin_dll_equilibrate',
    'thin_layer_dll_retry', 'thin_layer_begin', 'thin_layer_end',
    'accuracy_bound', 'audit', 'run_end', 'error_origin',
}


def same(a, b):
    if isinstance(b, str):
        return str(a) == b
    try:
        return abs(float(a) - float(b)) <= 1e-12 * max(1., abs(float(b)))
    except (TypeError, ValueError):
        return False


def sha256(path):
    value = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            value.update(block)
    return value.hexdigest()

CASE = json.loads(r'''{
 "case": "P02A_THIN",
 "source_case": {
  "id": "P02A",
  "title": "Jia2024 / Cheng homogeneous Case 2, psi=0, associated twin",
  "outline": [
   [
    0.0,
    0.0
   ],
   [
    20.0,
    0.0
   ],
   [
    20.0,
    4.0
   ],
   [
    16.0,
    4.0
   ],
   [
    10.0,
    10.0
   ],
   [
    0.0,
    10.0
   ]
  ],
  "materials": [
   {
    "name": "soil",
    "E": 14000.0,
    "nu": 0.3,
    "gamma": 20.0,
    "c": 10.0,
    "phi": 25.0,
    "psi": 25.0
   }
  ],
  "regions": [
   {
    "name": "soil",
    "xy": [
     [
      0.0,
      0.0
     ],
     [
      20.0,
      0.0
     ],
     [
      20.0,
      4.0
     ],
     [
      16.0,
      4.0
     ],
     [
      10.0,
      10.0
     ],
     [
      0.0,
      10.0
     ]
    ]
   }
  ],
  "h": [
   1.25,
   0.625,
   0.3125
  ],
  "policy": "equal_reduced_phi",
  "area": 158.0,
  "surcharge": false,
  "pore": false
 },
 "points": [
  [
   1,
   0.0,
   0.0
  ],
  [
   2,
   20.0,
   0.0
  ],
  [
   3,
   20.0,
   4.0
  ],
  [
   4,
   16.0,
   4.0
  ],
  [
   5,
   10.0,
   10.0
  ],
  [
   6,
   0.0,
   10.0
  ]
 ],
 "regions": [
  [
   1,
   1,
   0,
   "1,2,3,4,5,6"
  ]
 ],
 "materials": [
  [
   1,
   10.0,
   25.0,
   20.0,
   25.0
  ]
 ],
 "faces": [
  [
   1,
   1,
   1,
   2,
   0,
   "固定XY",
   0,
   0.0,
   0.0,
   0.0,
   0.0
  ],
  [
   2,
   1,
   2,
   3,
   0,
   "固定X",
   0,
   0.0,
   0.0,
   0.0,
   0.0
  ],
  [
   3,
   1,
   3,
   4,
   0,
   "自由",
   0,
   0.0,
   0.0,
   0.0,
   0.0
  ],
  [
   4,
   1,
   4,
   5,
   0,
   "自由",
   0,
   0.0,
   0.0,
   0.0,
   0.0
  ],
  [
   5,
   1,
   5,
   6,
   0,
   "自由",
   0,
   0.0,
   0.0,
   0.0,
   0.0
  ],
  [
   6,
   1,
   6,
   1,
   1,
   "固定X",
   0,
   0.0,
   0.0,
   0.0,
   0.0
  ]
 ],
 "boundary_explicit": [
  {
   "a": [
    0.0,
    0.0
   ],
   "b": [
    20.0,
    0.0
   ],
   "kind": "固定XY",
   "load0": [
    0.0,
    0.0
   ],
   "load1": [
    0,
    0
   ],
   "region": 1
  },
  {
   "a": [
    20.0,
    0.0
   ],
   "b": [
    20.0,
    4.0
   ],
   "kind": "固定X",
   "load0": [
    0.0,
    0.0
   ],
   "load1": [
    0,
    0
   ],
   "region": 1
  },
  {
   "a": [
    20.0,
    4.0
   ],
   "b": [
    16.0,
    4.0
   ],
   "kind": "自由",
   "load0": [
    0.0,
    0.0
   ],
   "load1": [
    0,
    0
   ],
   "region": 1
  },
  {
   "a": [
    16.0,
    4.0
   ],
   "b": [
    10.0,
    10.0
   ],
   "kind": "自由",
   "load0": [
    0.0,
    0.0
   ],
   "load1": [
    0,
    0
   ],
   "region": 1
  },
  {
   "a": [
    10.0,
    10.0
   ],
   "b": [
    0.0,
    10.0
   ],
   "kind": "自由",
   "load0": [
    0.0,
    0.0
   ],
   "load1": [
    0,
    0
   ],
   "region": 1
  },
  {
   "a": [
    0.0,
    10.0
   ],
   "b": [
    0.0,
    0.0
   ],
   "kind": "固定X",
   "load0": [
    0.0,
    0.0
   ],
   "load1": [
    0,
    0
   ],
   "region": 1
  }
 ],
 "settings": {
  "MODE": "Fs",
  "TARGET": 1536,
  "GRAVITY": "固定",
  "ADAPT": 1,
  "CYCLES": 3,
  "QUADRATURE": 8,
  "FS_TOL": 0.0001,
  "GAP": 0.01,
  "C0_AUTO": 0,
  "RAW_LOWER_AUDIT": 1,
  "TEST_INJECTION": 0,
  "HSD_CAPTURE": 0,
  "DIAGNOSTIC_MODE": 1,
  "ANALYSIS_POLICY": "ASSOCIATED",
  "SEISMIC_KH": 0.0,
  "SEISMIC_DIR": "右(+x)"
 },
 "region_areas": [
  158.0
 ],
 "total_weight": 3160.0,
 "scope_notes": [
  "G1R17 DLL5 full production profile: TARGET1536/adapt3/q8; one Excel at a time",
  "N cases use the unchanged Davis equivalent associated model",
  "E and nu are not inputs of the rigid-plastic model; D04A/D09A are still actually executed",
  "No new full VBA timing runs; G1R12 historic timings are not a controlled DLL-only comparison",
  "Fix8 forced thin-layer smoke: P02A geometry with ADAPT_THIN_LAYER=2 and GAP 0.005 so the chains run; functional test of the new code path, not a regression or speed case"
 ],
 "water": {
  "coupling": 0,
  "mode": 2,
  "gamma_w": 9.81,
  "suction": 0,
  "water_table": [],
  "seep_target": 1536
 },
 "extra_settings": {
  "GAP": 0.005,
  "ADAPT_THIN_LAYER": 2,
  "ADAPT_THIN_LAYER_MID_RATIO": 0.5,
  "ADAPT_THIN_UPPER_ROUNDS": 3,
  "ADAPT_THIN_LOWER_ROUNDS": 2,
  "ADAPT_THIN_LOWER_START": 1,
  "ADAPT_THIN_UPPER_CAP": 3000,
  "ADAPT_THIN_LOWER_CAP": 3000
 }
}''')

def settings(book):
    return {str(r[3]):r[1] for r in book.Worksheets('設定').Range('A4:D400').Value2 if r[3]}
SEISMIC_DEFAULTS={'SEISMIC_KH':0.0,'SEISMIC_DIR':'右(+x)'} # RPX_Seismic.bas RPX_Setting defaults when the row is absent
def configure(book,rec,call):
    ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
    for first,last,name in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
        data=tuple(tuple(row) for row in rec[name]);rng=ws.Range(first+'7:'+last+str(6+len(data)))
        rng.Value2=data;assert rng.Value2==data,(rec['case'],name)
    ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents()
    data=tuple(tuple(row) for row in rec['materials']);rng=ws.Range('A5:E'+str(4+len(data)))
    rng.Value2=data;assert rng.Value2==data
    # Product preset first (it also ensures missing setting rows); then only case-specific keys from the input.
    # Preset-owned keys from the G1R17 inputs (TARGET 1536, CYCLES 3, GAP 0.01) are never written.
    call(PRESET_MACRO)
    ws=book.Worksheets('設定');keys={str(row[3]):4+i for i,row in enumerate(ws.Range('A4:D400').Value2) if row[3]}
    absent={}
    for key in CASE_KEYS:
        if key not in keys and key in SEISMIC_DEFAULTS and rec['settings'][key]==SEISMIC_DEFAULTS[key]:
            absent[key]='row absent in delivery workbook; code default %r equals input'%(SEISMIC_DEFAULTS[key],);continue
        assert key in keys,('SETTING_MISSING',key)
        ws.Cells(keys[key],2).Value2=rec['settings'][key]
    # Fix8: optional extra settings (forced smoke variants only), written into free rows 4..120.
    for key,value in rec.get('extra_settings',{}).items():
        keys={str(row[3]):4+i for i,row in enumerate(ws.Range('A4:D120').Value2) if row[3]}
        if key not in keys:
            free=[4+i for i,row in enumerate(ws.Range('A4:D120').Value2) if not row[3] and not row[1]]
            assert free,('NO_FREE_SETTING_ROW',key);row=free[0];ws.Cells(row,4).Value2=key;keys[key]=row
        ws.Cells(keys[key],2).Value2=value
    actual=settings(book);assert all(actual[k]==rec['settings'][k] for k in CASE_KEYS if k not in absent)
    assert all(same(actual.get(k),v) for k,v in rec.get('extra_settings',{}).items()),('EXTRA_SETTINGS_NOT_APPLIED',rec.get('extra_settings'))
    actual['_absent_case_keys_using_default']=absent
    water=rec['water'];ws=book.Worksheets('浸透入力');assert ws.Range('Z1').Value2=='RPX_SEEP_LAYOUT_V2'
    ws.Range('A7:C206,E7:K206,M7:N206,P7:Q206').ClearContents()
    ws.Range('B3').Value2=ws.Range('Z6').Value2 # static groundwater level, mode=2
    ws.Range('B4').Value2=water['gamma_w'];ws.Range('D4').Value2=water['coupling'];ws.Range('F4').Value2=water['suction']
    ws.Range('C209').Value2=water['seep_target']
    if water['water_table']:ws.Range('M7:N8').Value2=tuple(tuple(row) for row in water['water_table'])
    for i,m in enumerate(rec['materials'],7):ws.Range(f'A{i}:C{i}').Value2=((m[0],1.,1.),)
    if water['coupling']:
        assert ws.Range('M7:N8').Value2==tuple(tuple(row) for row in water['water_table'])
    assert ws.Range('D4').Value2==water['coupling']
    # The delivery template contains levee seepage output. Dry cases must not
    # display or retain that unrelated field; wet cases regenerate their field.
    book.Worksheets('浸透結果').Cells.ClearContents()
    book.Worksheets('解析結果').Range('A1').Value2='未解析 / '+rec['case']
    for sheet in book.Worksheets:
        if '実行案内' in sheet.Name:sheet.Range('A1').Value2=rec['case']+' / G1R20 Fix9 rc8候補 動作テスト（2%標準プリセット）。入力表がこのケースの原本。'
    return actual
def read_log(path):
    with path.open(encoding='utf-16', newline='') as stream:
        return list(csv.DictReader(stream, delimiter='\t'))
def parts(value):return dict(p.split('=',1) for p in value.split(';') if '=' in p)

def utc():
    return datetime.now(timezone.utc).isoformat()


def write_json(path, data):
    def safe(value):
        if isinstance(value, float) and not math.isfinite(value):
            return str(value)
        if isinstance(value, dict):
            return {k: safe(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [safe(v) for v in value]
        return value

    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(safe(data), ensure_ascii=False, indent=2,
                               allow_nan=False) + "\n", encoding="utf-8")
    temp.replace(path)


def excel_processes():
    return [p.info["pid"] for p in psutil.process_iter(["pid", "name"])
            if (p.info["name"] or "").lower() == "excel.exe"]


def refuse_existing_excel():
    pids = excel_processes()
    if pids:
        raise RuntimeError(f"Close all existing Excel processes before running: {pids}")


def set_settings(book, values):
    """Same key/value columns and free-row policy as runner.configure."""
    ws = book.Worksheets("設定")
    for key, value in values.items():
        rows = ws.Range("A4:D400").Value2
        keys = {str(row[3]): 4 + i for i, row in enumerate(rows) if row[3]}
        if key not in keys:
            free = [4 + i for i, row in enumerate(rows[:117])
                    if not row[3] and not row[1]]  # rows 4..120
            if not free:
                raise RuntimeError(f"NO_FREE_SETTING_ROW: {key}")
            keys[key] = free[0]
            ws.Cells(keys[key], 4).Value2 = key
        ws.Cells(keys[key], 2).Value2 = value


def worker(workbook, folder, scenario):
    """All COM calls live here, so the parent can enforce a real wall deadline."""
    state = {"scenario": scenario, "started_utc": utc(), "phase": "STARTING"}
    app = book = pythoncom = None

    def persist(phase):
        state["phase"] = phase
        write_json(folder / "worker_state.json", state)

    # No pywin32 import or COM initialization in the supervisor or --help path.
    with (folder / "worker.log").open("w", encoding="utf-8", buffering=1) as log, \
            redirect_stdout(log), redirect_stderr(log):
        try:
            persist("STARTING")
            refuse_existing_excel()
            import pythoncom
            import win32com.client
            import win32process

            pythoncom.CoInitialize()
            persist("DISPATCHING")
            app = win32com.client.DispatchEx("Excel.Application")
            pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
            state["excel"] = {"pid": pid, "created": psutil.Process(pid).create_time()}
            persist("EXCEL_CREATED")  # publish ownership before any workbook/macro calls
            if set(excel_processes()) - {pid}:
                raise RuntimeError("Another Excel appeared during startup")
            app.Visible = False
            app.DisplayAlerts = False
            state["excel_version"] = str(app.Version)
            state["automation_security"] = app.AutomationSecurity
            book = app.Workbooks.Open(str(workbook), UpdateLinks=0)

            def call(name, *args):
                return app.Run("'" + book.Name.replace("'", "''") + "'!" + name, *args)

            rec = CASE
            baseline = configure(book, rec, call)  # preset -> case keys -> extra_settings
            absent = baseline.pop("_absent_case_keys_using_default")
            overrides = dict(SHORT)
            if scenario == 'E0':
                overrides['ACCURACY_DLL_UPPER_EQUILIBRATE'] = 0
            set_settings(book, overrides)
            actual = settings(book)
            expected = {**PRESET_EXPECTED,
                        **{k: rec["settings"][k] for k in CASE_KEYS if k not in absent},
                        **rec["extra_settings"], **overrides}
            mismatch = {k: {"expected": v, "actual": actual.get(k)}
                        for k, v in expected.items() if not same(actual.get(k), v)}
            # Also verify every pre-existing non-overridden setting stayed unchanged.
            mismatch.update({k: {"before": v, "after": actual.get(k)}
                             for k, v in baseline.items()
                             if k not in overrides and actual.get(k) != v
                             and not same(actual.get(k), v)})
            state.update(baseline_settings=baseline, actual_settings=actual,
                         short_overrides=overrides, settings_mismatch=mismatch,
                         absent_case_keys_using_default=absent)
            persist("CONFIGURED")
            if mismatch:
                raise RuntimeError(f"SETTINGS_MISMATCH: {mismatch}")
            book.Worksheets("解析結果").Range("B4:C4,B10:C10").ClearContents()
            call("RPX_NativeSetMode", "auto")
            state["native_info"] = str(call("RPX_NativeInfo"))
            if not state["native_info"].startswith("DLL_READY"):
                raise RuntimeError(state["native_info"])
            call("RPX_SeepSetSilent", True)
            persist("MESH")
            call("RPX_CreateMesh")
            # Keep all test hooks disarmed for both numerical scenarios.
            call("RPX_TestDisarm")
            book.Save()
            persist("RUNNING")
            started = time.monotonic()
            try:
                call("RPX_Run")  # synchronous; supervised by a separate process wall deadline
                state["run_returned"] = True
            except Exception:
                state["macro_error"] = traceback.format_exc()
            state["com_analysis_seconds"] = time.monotonic() - started
            ws = book.Worksheets("解析結果")
            state["sheet_status"] = ws.Range("A1").Value2
            state["lower_fs"], state["upper_fs"] = ws.Range("B4:C4").Value2[0]
            state["element_counts"] = list(ws.Range("B10:C10").Value2[0])
            state["relative_gap"] = ws.Range("B9").Value2
            persist("COLLECTED")
            call("RPX_TestDisarm")
            book.Save()
        except Exception:
            state["worker_error"] = traceback.format_exc()
        finally:
            persist("CLOSING")
            for obj, method, args in ((book, "Close", (False,)), (app, "Quit", ())):
                if obj is not None:
                    try:
                        getattr(obj, method)(*args)
                    except Exception:
                        state.setdefault("cleanup_errors", []).append(traceback.format_exc())
            book = app = None
            gc.collect()
            if pythoncom is not None:
                pythoncom.CoUninitialize()
            state["finished_utc"] = utc()
            persist("FINISHED")


def read_state(folder):
    try:
        return json.loads((folder / "worker_state.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def discover_startup_excel(started_epoch):
    """Cover a stuck DispatchEx/Hwnd before the worker can register its PID.

    The supervisor already required zero Excel processes. Adopt only a single
    new COM automation server, never an Excel opened interactively in the meantime.
    """
    candidates = []
    for proc in psutil.process_iter(["pid", "name", "create_time", "cmdline"]):
        info = proc.info
        flags = {arg.lower() for arg in (info["cmdline"] or [])}
        if ((info["name"] or "").lower() == "excel.exe"
                and info["create_time"] is not None and info["create_time"] >= started_epoch
                and flags & {"/automation", "-embedding", "/embedding"}):
            candidates.append({"pid": info["pid"], "created": info["create_time"]})
    return candidates[0] if len(candidates) == 1 else None


def kill_owned_excel(state, grace=0):
    """Never use taskkill /IM: verify BOTH the registered PID and creation time."""
    owner = state.get("excel")
    if not owner:
        return {"action": "no_registered_excel"}
    try:
        proc = psutil.Process(owner["pid"])
        if (proc.name().lower() != "excel.exe" or
                abs(proc.create_time() - owner["created"]) > 0.01):
            return {"action": "identity_mismatch", **owner}
        if grace:
            try:
                proc.wait(timeout=grace)
                return {"action": "exited", **owner}
            except psutil.TimeoutExpired:
                pass
        proc.kill()
        proc.wait(timeout=5)
        return {"action": "killed", **owner}
    except psutil.NoSuchProcess:
        return {"action": "exited", **owner}
    except psutil.Error as exc:
        return {"action": "cleanup_error", "error": str(exc), **owner}


def number(value, default=None):
    try:
        result = float(value)
        return result if math.isfinite(result) else default
    except (TypeError, ValueError):
        return default


def integer(value, default=None):
    result = number(value)
    return int(result) if result is not None and result == int(result) else default


def truth(value):
    return str(value).lower() in ('true', '1', '-1')


def finite(value):
    return type(value) in (int, float) and math.isfinite(value) and value > 0


def log_checks(scenario, state, rows):
    """Attribute diagnostics by native call ID and by enclosing direct trial.

    A native status is diagnostic only. Certified run_end and endpoint checks
    below remain separate from native result statuses.
    """
    details = [(r.get('event'), parts(r.get('detail') or '')) for r in rows]
    calls = [d for event, d in details if event == 'accuracy_dll_call']
    call_ids = [integer(d.get('call_id'), 0) for d in calls]
    by_id = {integer(d.get('call_id'), 0): d for d in calls}
    errors = []
    expected_setting = 1 if scenario == 'E1' else 0
    actual = state.get('actual_settings', {})
    upper = lower = helper = 0
    for call in calls:
        cid = integer(call.get('call_id'), 0)
        required = {'call_id', 'program_mode', 'program_role', 'original_primal',
                    'upper_equilibrate_setting', 'thin_original', 'thin_upper_retry',
                    'requested_seconds', 'packed_seconds', 'iteration_limit', 'equilibrate',
                    'state', 'poll_status', 'result_status', 'iterations', 'cancelled',
                    'finite_candidate', 'status_is_not_certificate'}
        if not required <= call.keys():
            errors.append(f'call {cid}: missing schema fields {sorted(required - call.keys())}')
        if call.get('status_is_not_certificate') != '1':
            errors.append(f'call {cid}: missing status-is-not-certificate attribution')
        # A thin chain can exhaust its budget before settings are packed. Such
        # attempts retain sentinel -1; they cannot demonstrate selector coverage.
        if integer(call.get('equilibrate')) == -1:
            if truth(call.get('finite_candidate')) or integer(call.get('upper_equilibrate_setting')) != -1:
                errors.append(f'call {cid}: unconfigured attempt claims a candidate/selector')
            continue
        role, mode = call.get('program_role'), call.get('program_mode')
        original = truth(call.get('original_primal'))
        scoped = role == 'original_primal' and mode == 'upper' and original
        if scoped:
            upper += 1
        elif role == 'original_primal' and mode == 'lower' and original:
            lower += 1
        if role == 'original_dual_helper':
            helper += 1
        if integer(call.get('upper_equilibrate_setting')) != (expected_setting if scoped else -1):
            errors.append(f'call {cid}: upper selector applied outside literal upper or wrong scenario')
        if role not in ('original_primal', 'original_dual_helper', 'auxiliary'):
            errors.append(f'call {cid}: missing/unknown program_role')
        if original != (role == 'original_primal'):
            errors.append(f'call {cid}: inconsistent original_primal/program_role')
        # These are the rc9a branches, with only scoped E1 overriding them.
        if scoped and expected_setting == 1:
            expected_equ = 1
        elif truth(call.get('thin_original')):
            expected_equ = integer(actual.get('ADAPT_THIN_DLL_EQUILIBRATE', 1))
        elif truth(call.get('thin_upper_retry')):
            expected_equ = integer(actual.get('ADAPT_THIN_UPPER_DLL_RETRY_EQUILIBRATE', 1))
        elif mode in ('upper', 'lower'):
            expected_equ = integer(actual.get('ACCURACY_DLL_EQUILIBRATE', 0))
        else:
            expected_equ = integer(actual.get('ACCURACY_DLL_DUAL_EQUILIBRATE', 1))
        if integer(call.get('equilibrate')) != expected_equ:
            errors.append(f'call {cid}: packed equilibrate disagrees with applicable selector')
    vectors = [d for event, d in details if event == 'accuracy_dual_vector']
    for vector in vectors:
        cid = integer(vector.get('call_id'), 0)
        required = {'call_id', 'candidate_role', 'mode', 'solver_status', 'cost_inf',
                    'cost_index', 'dual_inf', 'aty_inf_derived', 'aty_inf_derived_valid',
                    'stat_index', 'stat_at_cost_index', 'cost_at_cost_index', 'raw_stat',
                    'raw_cone', 'original_objective', 'accepted', 'physical_bound'}
        if not required <= vector.keys():
            errors.append(f'dual vector {cid}: missing fields {sorted(required - vector.keys())}')
        if cid <= 0 or cid not in by_id or not vector.get('candidate_role'):
            errors.append(f'dual vector: missing native origin or candidate_role ({cid})')
        if integer(vector.get('solver_status')) != integer(by_id.get(cid, {}).get('poll_status')):
            errors.append(f'dual vector {cid}: status disagrees with native origin')
        if vector.get('candidate_role') == 'helper_primal' and by_id.get(cid, {}).get('program_role') != 'original_dual_helper':
            errors.append(f'dual vector {cid}: helper vector points to a different program role')
    trial_calls, trial_vectors, trial_audits = set(), [], []
    trial_q = None
    trial_open = False
    ends = []
    unaudited = audited = 0
    for event, detail in details:
        if event == 'direct_dll_begin':
            if trial_open:
                errors.append('direct begin without preceding end')
            trial_open = True
            trial_calls, trial_vectors, trial_audits = set(), [], []
            trial_q = integer(detail.get('q'))
        elif event == 'accuracy_dll_call':
            trial_calls.add(integer(detail.get('call_id'), 0))
        elif event == 'accuracy_dual_vector':
            trial_vectors.append(detail)
        elif event == 'audit' and trial_open:
            trial_audits.append(detail)
        elif event == 'direct_dll_end':
            if not trial_open:
                errors.append('direct end without current begin')
            trial_open = False
            ends.append(detail)
            pcid = integer(detail.get('primal_call_id'), 0)
            hcid = integer(detail.get('helper_call_id'), 0)
            ccid = integer(detail.get('candidate_call_id'), 0)
            for field, cid in (('primal_call_id', pcid), ('helper_call_id', hcid),
                               ('candidate_call_id', ccid),
                               ('dual_call_id', integer(detail.get('dual_call_id'), 0))):
                if cid and cid not in trial_calls:
                    errors.append(f'direct end {field}={cid}: origin outside current trial')
            if pcid:
                primal_call = by_id.get(pcid, {})
                if integer(detail.get('primal_status')) != integer(primal_call.get('poll_status')):
                    errors.append(f'direct end primal call {pcid}: status mismatch')
                if integer(detail.get('primal_equilibrate')) != integer(primal_call.get('equilibrate')):
                    errors.append(f'direct end primal call {pcid}: equilibrate mismatch')
            if hcid and by_id.get(hcid, {}).get('program_role') != 'original_dual_helper':
                errors.append(f'direct end helper call {hcid}: role mismatch')
            if detail.get('solver_status') != detail.get('primal_status'):
                errors.append('direct end: solver_status does not describe primal phase')
            for legacy, explicit in (('raw_eq', 'primal_raw_eq'), ('raw_cone', 'primal_raw_cone')):
                if not same(detail.get(legacy), number(detail.get(explicit))):
                    errors.append(f'direct end: {legacy} mixes primal and dual phases')
            if not truth(detail.get('primal_audited')):
                if (truth(detail.get('primal_audit_pass'))
                        or any(number(detail.get(k)) != 0 for k in
                               ('primal_raw_eq', 'primal_raw_cone', 'primal_objective'))):
                    errors.append('direct end: unaudited primal fields were not reset')
            if truth(detail.get('dual_audited')):
                audited += 1
                dcid = integer(detail.get('dual_call_id'), 0)
                role = detail.get('dual_role')
                candidates = [v for v in trial_vectors if integer(v.get('call_id'), 0) == dcid
                              and v.get('candidate_role') == role]
                if not candidates:
                    errors.append(f'direct end: dual origin {dcid}/{role} has no audit in current trial')
                else:
                    vector = candidates[-1]
                    if detail.get('dual_status') != vector.get('solver_status'):
                        errors.append(f'direct end: stale dual_status for {dcid}/{role}')
                    for final_key, vector_key in (('dual_stat', 'raw_stat'), ('dual_cone', 'raw_cone'),
                                                   ('dual_objective', 'original_objective')):
                        if not same(detail.get(final_key), number(vector.get(vector_key))):
                            errors.append(f'direct end: stale {final_key} for {dcid}/{role}')
            else:
                unaudited += 1
                if (integer(detail.get('dual_call_id'), -1) != 0
                        or integer(detail.get('dual_status'), -1) != 0 or detail.get('dual_role') != ''
                        or any(number(detail.get(k)) != 0 for k in
                               ('dual_stat', 'dual_cone', 'dual_objective'))):
                    errors.append('direct end: unaudited dual fields were not reset')
            if detail.get('kind') == 'original_primal' and truth(detail.get('accepted')):
                if (not truth(detail.get('primal_audited'))
                        or not truth(detail.get('primal_audit_pass')) or not pcid or ccid != pcid):
                    errors.append('accepted original primal: missing audit or wrong candidate origin')
                if trial_q != 8 or any(number(detail.get(k), math.inf) > 0.0000001
                                       for k in ('primal_raw_eq', 'primal_raw_cone')):
                    errors.append('accepted original primal: q8/raw residual audit evidence missing')
                # FieldGate and final commit both emit existing physical audits.
                # Use the last occurrence per category so a rejected earlier
                # normalized/recenter attempt does not mask the accepted field.
                physical = {a.get('category'): a for a in trial_audits
                            if a.get('mode') == detail.get('mode')}
                if not physical or any(a.get('status') != 'PASS' or not truth(a.get('category_pass'))
                                       or number(a.get('max_raw'), math.inf) > number(a.get('threshold'), -math.inf)
                                       for a in physical.values()):
                    errors.append('accepted original primal: no complete passing physical audit log')
            if detail.get('kind') == 'original_dual_helper' and truth(detail.get('accepted')):
                if by_id.get(ccid, {}).get('program_role') != 'original_dual_helper':
                    errors.append('accepted helper: wrong candidate origin')
            last_id = integer(detail.get('last_call_id'), 0)
            if not last_id or any(cid > last_id for cid in trial_calls):
                errors.append('direct end: invalid last_call_id')
    if trial_open:
        errors.append('direct trial did not end')
    accepted_original = 0
    for index, row in enumerate(rows):
        detail = parts(row.get('detail') or '')
        if (row.get('event') != 'direct_dll_end' or detail.get('kind') != 'original_primal'
                or not truth(detail.get('accepted'))):
            continue
        accepted_original += 1
        sid = integer(row.get('solve_id'), 0)
        following = []
        for later in rows[index + 1:]:
            if integer(later.get('solve_id'), 0) != sid:
                break
            following.append(later)
            if later.get('event') == 'solve_end':
                break
        finish = following[-1] if following and following[-1].get('event') == 'solve_end' else None
        expected_status = ('ORIGINAL_ACCURACY_UPPER_WITNESS' if detail.get('mode') == 'upper'
                           else 'ORIGINAL_FIXED_LOWER_WITNESS')
        if (sid <= 0 or not finish or parts(finish.get('detail') or '').get('status') != expected_status
                or any(r.get('event') in ('hsd_start', 'solve_start') for r in following)):
            errors.append(f'accepted original primal solve {sid}: missing witness end or unexpected HSD continuation')
    checks = {
        'native_calls_present': bool(calls),
        'call_ids_positive_unique_increasing': bool(call_ids) and all(x > 0 for x in call_ids)
            and len(set(call_ids)) == len(call_ids) and call_ids == sorted(call_ids),
        'no_attribution_or_equilibrate_errors': not errors,
        'direct_end_rows_present': bool(ends),
        'no_early_reject_events': not any('early_reject' in str(event) for event, _ in details),
        'no_test_injection': not any(event in ('test_injection', 'test_no_iterate')
                                     for event, _ in details),
    }
    coverage = {
        'upper_original_primal_observed': upper > 0,
        'lower_original_primal_observed': lower > 0,
        'dual_helper_observed': helper > 0,
        'dual_vector_observed': bool(vectors),
        'audited_dual_end_observed': audited > 0,
        'unaudited_dual_end_observed': unaudited > 0,
        'thin_layer_begin_observed': any(event == 'thin_layer_begin' for event, _ in details),
    }
    return checks, coverage, {
        'errors': errors, 'native_call_count': len(calls), 'upper_original_count': upper,
        'lower_original_count': lower, 'helper_count': helper, 'dual_vector_count': len(vectors),
        'direct_end_count': len(ends), 'audited_dual_end_count': audited,
        'unaudited_dual_end_count': unaudited, 'accepted_original_primal_count': accepted_original,
    }


def grade(scenario, state, rows, *, timed_out, exitcode, log_error=None):
    ends = [parts(row.get('detail') or '') for row in rows if row.get('event') == 'run_end']
    terminal = ends[0].get('status') if len(ends) == 1 else None
    lower, upper = state.get('lower_fs'), state.get('upper_fs')
    checks = {
        'worker_finished': state.get('phase') == 'FINISHED' and exitcode == 0,
        'within_wall_limit': not timed_out,
        'no_driver_errors': not any(state.get(k) for k in ('worker_error', 'macro_error', 'cleanup_errors')),
        'settings_verified': state.get('settings_mismatch') == {},
        'run_returned': state.get('run_returned') is True,
        'analysis_log_read': log_error is None and bool(rows),
        'one_run_end': len(ends) == 1,
        'finite_ordered_fs': finite(lower) and finite(upper) and lower <= upper,
        'certified_run_end': terminal in ('OPTIMAL', 'PARTIAL_CERTIFIED'),
    }
    logging, coverage, evidence = log_checks(scenario, state, rows)
    checks.update(logging)
    passed = all(checks.values())
    status = ('FAIL' if not passed else 'PASS' if all(coverage.values()) else 'INCOMPLETE')
    return {**state, 'scenario': scenario, 'case': CASE['case'], 'status': status,
            'run_end_status': terminal, 'checks': checks,
            'failed_checks': [k for k, value in checks.items() if not value],
            'coverage': coverage, 'unobserved_coverage': [k for k, value in coverage.items() if not value],
            'log_evidence': evidence,
            'events': [row for row in rows if row.get('event') in EVENTS],
            'log_error': log_error, 'timed_out': timed_out, 'worker_exitcode': exitcode}


def run_scenario(source, out, scenario, limit):
    refuse_existing_excel()
    folder = out / f"scenario_{scenario}"
    folder.mkdir()
    workbook = folder / f"repro_rc9b_{scenario}.xlsm"
    shutil.copy2(source, workbook)
    shutil.copy2(source.with_name("rpfem_native_x64.dll"), folder / "rpfem_native_x64.dll")
    copied_dll_sha = sha256(folder / 'rpfem_native_x64.dll')
    proc = mp.get_context("spawn").Process(target=worker, args=(workbook, folder, scenario))
    started = time.monotonic()
    started_epoch = time.time()
    proc.start()
    interrupted = False
    monitor_error = None
    startup_owner = None
    try:
        while proc.is_alive() and time.monotonic() - started < limit:
            state = read_state(folder)
            if state.get("phase") == "DISPATCHING" and not startup_owner:
                startup_owner = discover_startup_excel(started_epoch)
                if startup_owner:
                    write_json(folder / "startup_excel.json", startup_owner)
            proc.join(timeout=min(0.5, max(0, limit - (time.monotonic() - started))))
    except KeyboardInterrupt:
        interrupted = True
    except Exception:
        # Monitoring errors must still pass through process cleanup.
        interrupted = True
        monitor_error = traceback.format_exc()
    timed_out = proc.is_alive() and not interrupted
    state = read_state(folder)
    if not state.get("excel"):
        startup_owner = startup_owner or discover_startup_excel(started_epoch)
        if startup_owner:
            state["excel"] = startup_owner
    cleanup = kill_owned_excel(state, grace=0 if proc.is_alive() else 3)
    if proc.is_alive():
        proc.terminate()
        proc.join(timeout=5)
        if proc.is_alive():
            proc.kill()
            proc.join(timeout=5)
    state = read_state(folder)
    if interrupted:
        state["worker_error"] = monitor_error or "Interrupted by operator"
    if cleanup["action"] in ("identity_mismatch", "cleanup_error"):
        state.setdefault("cleanup_errors", []).append(str(cleanup))
    rows, log_error, log_path = [], None, None
    try:
        logs = sorted((folder / "RPFEM_logs").glob("*_analysis*.tsv"))
        if len(logs) != 1:
            raise RuntimeError(f"Expected one analysis log, found {len(logs)}: {logs}")
        log_path = folder / "analysis.tsv"
        shutil.copy2(logs[0], log_path)
        rows = read_log(log_path)  # original UTF-16 TSV reader
    except Exception:
        log_error = traceback.format_exc()
    result = grade(scenario, state, rows, timed_out=timed_out,
                   exitcode=proc.exitcode, log_error=log_error)
    result.update(workbook=str(workbook), analysis_log=str(log_path) if log_path else None,
                  excel_cleanup=cleanup, wall_seconds=time.monotonic() - started,
                  wall_limit_seconds=limit)
    result['copied_dll_unchanged'] = copied_dll_sha == sha256(folder / 'rpfem_native_x64.dll')
    result['checks']['copied_dll_unchanged'] = result['copied_dll_unchanged']
    if not result['copied_dll_unchanged']:
        result['status'] = 'FAIL'
        result['failed_checks'].append('copied_dll_unchanged')
    write_json(folder / "result.json", result)
    return result, interrupted


def compare_fs(candidate, baseline, tolerance, *, equal=False):
    keys = ('lower_fs', 'upper_fs')
    valid = all(finite(row.get(k)) for row in (candidate, baseline) for k in keys)
    if not valid:
        return {'status': 'FAIL', 'reason': 'Missing finite certified endpoint pair'}
    lower, upper = (candidate[k] for k in keys)
    base_lower, base_upper = (baseline[k] for k in keys)
    width = (upper - lower) / upper
    base_width = (base_upper - base_lower) / base_upper
    checks = {
        'candidate_certified': candidate.get('run_end_status') in ('OPTIMAL', 'PARTIAL_CERTIFIED'),
        'baseline_certified': baseline.get('run_end_status') in ('OPTIMAL', 'PARTIAL_CERTIFIED'),
        'lower_same_or_better': lower >= base_lower - tolerance,
        'upper_same_or_better': upper <= base_upper + tolerance,
        'relative_interval_same_or_better': width <= base_width + tolerance,
    }
    if equal:
        checks['lower_reproduced'] = abs(lower - base_lower) <= tolerance
        checks['upper_reproduced'] = abs(upper - base_upper) <= tolerance
    return {'status': 'PASS' if all(checks.values()) else 'FAIL', 'checks': checks,
            'absolute_fs_tolerance': tolerance, 'relative_interval_tolerance': tolerance,
            'candidate_fs': [lower, upper], 'baseline_fs': [base_lower, base_upper],
            'candidate_relative_interval': width, 'baseline_relative_interval': base_width,
            'lower_delta': lower - base_lower, 'upper_delta': upper - base_upper}


def compare_settings(candidate, baseline):
    """The new upper selector is the only intended configuration difference."""
    left, right = candidate.get('actual_settings', {}), baseline.get('actual_settings', {})
    missing = object()
    differences = {}
    for key in sorted((left.keys() | right.keys()) - {'ACCURACY_DLL_UPPER_EQUILIBRATE'}):
        a, b = left.get(key, missing), right.get(key, missing)
        if a is missing and key in SEISMIC_DEFAULTS:
            a = SEISMIC_DEFAULTS[key]
        if b is missing and key in SEISMIC_DEFAULTS:
            b = SEISMIC_DEFAULTS[key]
        if a is missing or b is missing or not (a == b or same(a, b)):
            differences[key] = {'candidate': '<row absent>' if a is missing else a,
                                'baseline': '<row absent>' if b is missing else b}
    return {'status': 'PASS' if left and right and not differences else 'FAIL',
            'differences': differences, 'excluded_setting': 'ACCURACY_DLL_UPPER_EQUILIBRATE',
            'absent_row_defaults_allowed': SEISMIC_DEFAULTS}


def compare_run(candidate, baseline, tolerance, *, equal=False):
    result = compare_fs(candidate, baseline, tolerance, equal=equal)
    result['configuration'] = compare_settings(candidate, baseline)
    if result['configuration']['status'] != 'PASS':
        result['status'] = 'FAIL'
    return result


def load_rc9a_baseline(path):
    """Accept the disarmed scenario 0 of the supplied rc9a driver, or its result.

    The operator identifies this artifact as rc9a with --rc9a-baseline. We verify
    its completed/certified result and the same P02A_THIN short settings; we do
    not infer rc9a provenance from a filename or silently use historical Fs.
    """
    artifact = json.loads(path.read_text(encoding='utf-8-sig'))
    result = artifact.get('scenarios', {}).get('0') if 'scenarios' in artifact else artifact
    if not isinstance(result, dict) or result.get('scenario') != '0':
        raise ValueError('rc9a baseline must be a disarmed scenario 0 result/summary')
    if result.get('status') != 'PASS' or result.get('run_end_status') not in ('OPTIMAL', 'PARTIAL_CERTIFIED'):
        raise ValueError('rc9a baseline must have PASS and certified run_end')
    if result.get('settings_mismatch') != {}:
        raise ValueError('rc9a baseline settings were not verified')
    actual = result.get('actual_settings', {})
    expected = {**PRESET_EXPECTED,
                **{k: CASE['settings'][k] for k in CASE_KEYS},
                **CASE['extra_settings'], **SHORT}
    expected.pop('ACCURACY_DLL_UPPER_EQUILIBRATE')
    absent = result.get('absent_case_keys_using_default', {})
    mismatch = {key: {'expected': value, 'actual': actual.get(key)}
                for key, value in expected.items()
                if not (key in absent and key in SEISMIC_DEFAULTS
                        and same(value, SEISMIC_DEFAULTS[key]))
                and not same(actual.get(key), value)}
    if mismatch:
        raise ValueError(f'rc9a baseline is not the same short configuration: {mismatch}')
    input_name = Path(str(artifact.get('input', ''))).name
    if input_name and input_name != 'P02A_THIN.json':
        raise ValueError('rc9a baseline summary input is not P02A_THIN.json')
    if not all(finite(result.get(k)) for k in ('lower_fs', 'upper_fs')):
        raise ValueError('rc9a baseline has no finite endpoint pair')
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, required=True,
                        help='Built rc9b .xlsm; source workbook is not modified')
    parser.add_argument('--out', type=Path, required=True,
                        help='New/empty subdirectory of this driver folder')
    parser.add_argument('--scenario', choices=('all', 'E1', 'E0'), default='all',
                        help='Default: E1 preset defaults, then E0 legacy-selector control')
    parser.add_argument('--wall-limit', type=float, default=1500, metavar='SECONDS',
                        help='Combined wall budget including startup/mesh/cleanup (default: 1500)')
    parser.add_argument('--fs-tolerance', type=float, default=0.0001,
                        help='Absolute endpoint/relative interval comparison tolerance (default: 0.0001)')
    parser.add_argument('--rc9a-baseline', type=Path,
                        help='Optional archived rc9a scenario 0 result.json or repro_summary.json')
    args = parser.parse_args(argv)
    source, out = args.workbook.resolve(), args.out.resolve()
    if not math.isfinite(args.wall_limit) or args.wall_limit <= 60:
        parser.error('--wall-limit must be finite and greater than 60 seconds')
    if not math.isfinite(args.fs_tolerance) or args.fs_tolerance < 0:
        parser.error('--fs-tolerance must be finite and nonnegative')
    if not source.is_file() or source.suffix.lower() != '.xlsm':
        parser.error('--workbook must name an existing built rc9b .xlsm')
    dll = source.with_name('rpfem_native_x64.dll')
    if not dll.is_file():
        parser.error('Place rpfem_native_x64.dll beside the source workbook')
    if ROOT not in out.parents:
        parser.error('--out must be a subdirectory of this driver folder')
    if out.exists() and (not out.is_dir() or any(out.iterdir())):
        parser.error('--out must be new or empty; use a separate directory for each invocation')
    baseline = None
    if args.rc9a_baseline:
        try:
            baseline = load_rc9a_baseline(args.rc9a_baseline.resolve())
        except Exception as exc:
            parser.error(f'--rc9a-baseline: {exc}')
    refuse_existing_excel()
    scenarios = ('E1', 'E0') if args.scenario == 'all' else (args.scenario,)
    source_hashes = {'workbook': sha256(source), 'dll': sha256(dll)}
    out.mkdir(parents=True, exist_ok=True)
    summary = {
        'started_utc': utc(), 'case': CASE['case'], 'source_workbook': str(source),
        'source_sha256_before': source_hashes,
        'embedded_input_sha256': hashlib.sha256(json.dumps(CASE, ensure_ascii=False,
                                      sort_keys=True).encode('utf-8')).hexdigest(),
        'input_origin': '../codex_fix9rc9a_repro/inputs/P02A_THIN.json (embedded)',
        'preset_macro': PRESET_MACRO, 'wall_limit_seconds': args.wall_limit,
        'wall_limit_scope': 'combined; 30 seconds cleanup reserve per scenario',
        'fs_tolerance': args.fs_tolerance,
        'rc9a_baseline': str(args.rc9a_baseline.resolve()) if args.rc9a_baseline else None,
        'rc9a_baseline_sha256': sha256(args.rc9a_baseline.resolve()) if baseline else None,
        'rc9a_baseline_provenance': 'operator supplied; preset/case setup verified; full settings compared after run' if baseline else None,
        'status': 'RUNNING', 'scenarios': {},
        'rc9a_independent_comparison': {'status': 'UNVERIFIED', 'reason': 'No archived rc9a baseline supplied'},
    }
    summary_path = out / 'repro_summary.json'
    write_json(summary_path, summary)
    started = time.monotonic()
    deadline = started + args.wall_limit
    for index, scenario in enumerate(scenarios):
        remaining = len(scenarios) - index
        # The monitor hard-stops a synchronous/stuck macro before the global
        # deadline, leaving time for owned Excel and worker process cleanup.
        allowance = max(0, (deadline - time.monotonic() - 30 * remaining) / remaining)
        print(f'Scenario {scenario}: supervised work allowance {allowance:g}s', flush=True)
        try:
            if allowance <= 0:
                raise RuntimeError('Combined wall budget exhausted before scenario startup')
            result, interrupted = run_scenario(source, out, scenario, allowance)
        except Exception:
            result, interrupted = {'scenario': scenario, 'status': 'FAIL',
                                   'driver_error': traceback.format_exc()}, False
        summary['scenarios'][scenario] = result
        write_json(summary_path, summary)
        print(f"Scenario {scenario}: {result['status']}; run_end={result.get('run_end_status')}; "
              f"Fs={result.get('lower_fs')}, {result.get('upper_fs')}", flush=True)
        if interrupted:
            break
    completed = summary['scenarios']
    if {'E1', 'E0'} <= completed.keys():
        summary['E1_vs_E0'] = compare_run(completed['E1'], completed['E0'], args.fs_tolerance)
        summary['E1_vs_E0']['scope'] = 'rc9b E1 versus rc9b legacy-selector control; no independent rc9a provenance claim'
    else:
        summary['E1_vs_E0'] = {'status': 'UNVERIFIED', 'reason': 'Both E1 and E0 are required'}
    if baseline:
        comparisons = {scenario: compare_run(result, baseline, args.fs_tolerance, equal=scenario == 'E0')
                       for scenario, result in completed.items()}
        summary['rc9a_independent_comparison'] = {
            'status': 'PASS' if comparisons and all(c['status'] == 'PASS' for c in comparisons.values()) else 'FAIL',
            'scenarios': comparisons,
        }
    try:
        source_after = {'workbook': sha256(source), 'dll': sha256(dll)}
        summary['source_sha256_after'] = source_after
        summary['source_workbook_and_dll_unchanged'] = source_after == source_hashes
    except Exception:
        summary['source_workbook_and_dll_unchanged'] = False
        summary['source_integrity_error'] = traceback.format_exc()
    failures = (len(completed) != len(scenarios) or any(r['status'] == 'FAIL' for r in completed.values())
                or summary['E1_vs_E0']['status'] == 'FAIL'
                or summary['rc9a_independent_comparison']['status'] == 'FAIL'
                or not summary['source_workbook_and_dll_unchanged'])
    incomplete = any(r['status'] == 'INCOMPLETE' for r in completed.values())
    summary.update(status='FAIL' if failures else 'INCOMPLETE' if incomplete else 'PASS',
                   finished_utc=utc(), wall_seconds=time.monotonic() - started)
    summary['within_combined_wall_budget'] = summary['wall_seconds'] <= args.wall_limit
    if not summary['within_combined_wall_budget']:
        summary['status'] = 'FAIL'
    write_json(summary_path, summary)
    print(summary_path, flush=True)
    return 0 if summary['status'] == 'PASS' else 2 if summary['status'] == 'INCOMPLETE' else 1


if __name__ == '__main__':
    mp.freeze_support()
    raise SystemExit(main())


