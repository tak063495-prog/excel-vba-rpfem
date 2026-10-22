"""Operator-run rc9c midnight-wrap P02A_THIN reproduction. Import/--help never opens Excel.

The source workbook and adjacent DLL are read only; each scenario uses private
copies below --out. No module injection, build, or macro-security changes.
The embedded case/configuration comes from codex_fix9rc9a_repro. W0 uses the
normal clock; W1 forces the corrected fallback through a wrap at 120s.
"""

import argparse
import csv
import gc
import hashlib
import json
import math
import multiprocessing as mp
import re
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
# rc9b short numerical setup; identical test permission/case limit in W0/W1.
# Only the explicit 900s case deadline is added; solver gates stay unchanged.
SHORT = {
    'MODE': 'Fs', 'TARGET': 320, 'ADAPT': 1, 'CYCLES': 1,
    'TEST_INJECTION': 1, 'ADAPT_THIN_CASE_LIMIT_SECONDS': 900,
    'ACCURACY_DLL_RECOVERY': 1,
    'ADAPT_PRECISION_FROM_START': 1, 'ADAPT_THIN_LAYER': 2,
    'ADAPT_THIN_UPPER_ROUNDS': 1, 'ADAPT_THIN_UPPER_CAP': 800,
    'ADAPT_THIN_UPPER_ROUND_MAX_ELEMENTS': 800,
    'ADAPT_THIN_EARLY_UPPER_CAP': 800, 'ADAPT_THIN_LOWER_ROUNDS': 0,
    'ADAPT_THIN_LOWER_EXTRA_ROUNDS': 0, 'ADAPT_THIN_MID_SEPARATE': 0,
    'GAP': 0.000001, 'ADAPT_ACCURACY_CAP': 800,
    'ADAPT_ACCURACY_ROUNDS': 1, 'ADAPT_LOWER_TARGET': 0,
    'ADAPT_LOWER_FAN_ROUNDS': 0,
}
def same(a, b):
    if a is None and b is None:
        return True
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
def parts(value):
    return {key.strip(): item.strip() for key, item in
            (field.split('=', 1) for field in str(value or '').split(';') if '=' in field)}

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


WRAP_AT = 120.0
CLOCK_SLACK = 2.0  # accommodates Timer resolution, log rounding and COM overhead
CERTIFIED = {'OPTIMAL', 'PARTIAL_CERTIFIED'}


def worker(workbook, folder, scenario):
    """COM stays in a child; a synchronous/stuck macro has an external deadline."""
    state = {'scenario': scenario, 'started_utc': utc(), 'phase': 'STARTING'}
    app = book = pythoncom = None

    def persist(phase):
        state['phase'] = phase
        write_json(folder / 'worker_state.json', state)

    with (folder / 'worker.log').open('w', encoding='utf-8', buffering=1) as log, \
            redirect_stdout(log), redirect_stderr(log):
        try:
            persist('STARTING')
            refuse_existing_excel()
            import pythoncom
            import win32com.client
            import win32process

            pythoncom.CoInitialize()
            persist('DISPATCHING')
            app = win32com.client.DispatchEx('Excel.Application')
            pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
            state['excel'] = {'pid': pid, 'created': psutil.Process(pid).create_time()}
            persist('EXCEL_CREATED')
            if set(excel_processes()) - {pid}:
                raise RuntimeError('Another Excel appeared during startup')
            app.Visible = False
            app.DisplayAlerts = False
            state['excel_version'] = str(app.Version)
            state['automation_security'] = app.AutomationSecurity
            book = app.Workbooks.Open(str(workbook), UpdateLinks=0)

            def call(name, *args):
                return app.Run("'" + book.Name.replace("'", "''") + "'!" + name, *args)

            baseline = configure(book, CASE, call)
            absent = baseline.pop('_absent_case_keys_using_default')
            overrides = dict(SHORT)  # byte-for-byte identical settings for W0 and W1
            set_settings(book, overrides)
            actual = settings(book)
            expected = {**PRESET_EXPECTED,
                        **{k: CASE['settings'][k] for k in CASE_KEYS if k not in absent},
                        **CASE['extra_settings'], **overrides}
            mismatch = {k: {'expected': value, 'actual': actual.get(k)}
                        for k, value in expected.items() if not same(actual.get(k), value)}
            mismatch.update({k: {'before': value, 'after': actual.get(k)}
                             for k, value in baseline.items()
                             if k not in overrides and not same(actual.get(k), value)})
            state.update(baseline_settings=baseline, actual_settings=actual,
                         short_overrides=overrides, settings_mismatch=mismatch,
                         absent_case_keys_using_default=absent)
            persist('CONFIGURED')
            if mismatch:
                raise RuntimeError(f'SETTINGS_MISMATCH: {mismatch}')
            book.Worksheets('解析結果').Range('B4:C4,B10:C10').ClearContents()
            call('RPX_NativeSetMode', 'auto')
            state['native_info'] = str(call('RPX_NativeInfo'))
            if not state['native_info'].startswith('DLL_READY'):
                raise RuntimeError(state['native_info'])
            call('RPX_SeepSetSilent', True)
            persist('MESH')
            call('RPX_CreateMesh')
            call('RPX_TestDisarm')  # existing failure-injection hooks stay disarmed
            book.Save()
            # Arm only after configuration/mesh. There is no VBA project access.
            state['clock_before_arm'] = float(call('RPX_Clock'))
            arm_started = time.monotonic()
            call('RPX_TestClockWrap', WRAP_AT if scenario == 'W1' else 0.0)
            state['clock_hook_before'] = str(call('RPX_TestClockStatus'))
            state['clock_at_run_start'] = float(call('RPX_Clock'))
            state['arm_com_seconds'] = time.monotonic() - arm_started
            persist('RUNNING')
            started = time.monotonic()
            state['arm_to_macro_entry_seconds'] = started - arm_started
            try:
                call('RPX_Run')
                state['run_returned'] = True
            except Exception:
                state['macro_error'] = traceback.format_exc()
            state['com_analysis_seconds'] = time.monotonic() - started
            # Read before explicitly disarming. EndWork must retain clock-hook evidence.
            state['clock_at_run_end'] = float(call('RPX_Clock'))
            state['clock_hook_after'] = str(call('RPX_TestClockStatus'))
            ws = book.Worksheets('解析結果')
            state['sheet_status'] = ws.Range('A1').Value2
            state['lower_fs'], state['upper_fs'] = ws.Range('B4:C4').Value2[0]
            state['element_counts'] = list(ws.Range('B10:C10').Value2[0])
            state['relative_gap'] = ws.Range('B9').Value2
            persist('COLLECTED')
            call('RPX_TestClockWrap', 0.0)
            call('RPX_TestDisarm')
            book.Save()
        except Exception:
            state['worker_error'] = traceback.format_exc()
        finally:
            persist('CLOSING')
            for obj, method, args in ((book, 'Close', (False,)), (app, 'Quit', ())):
                if obj is not None:
                    try:
                        getattr(obj, method)(*args)
                    except Exception:
                        state.setdefault('cleanup_errors', []).append(traceback.format_exc())
            book = app = None
            gc.collect()
            if pythoncom is not None:
                pythoncom.CoUninitialize()
            state['finished_utc'] = utc()
            persist('FINISHED')


def read_state(folder):
    try:
        return json.loads((folder / 'worker_state.json').read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return {}


def discover_startup_excel(started_epoch):
    """Adopt only a single newly created COM server if DispatchEx/Hwnd is stuck."""
    candidates = []
    for proc in psutil.process_iter(['pid', 'name', 'create_time', 'cmdline']):
        info = proc.info
        flags = {arg.lower() for arg in (info['cmdline'] or [])}
        if ((info['name'] or '').lower() == 'excel.exe'
                and info['create_time'] is not None and info['create_time'] >= started_epoch
                and flags & {'/automation', '-embedding', '/embedding'}):
            candidates.append({'pid': info['pid'], 'created': info['create_time']})
    return candidates[0] if len(candidates) == 1 else None


def kill_owned_excel(state, grace=0):
    """Verify both PID and creation time; never terminate Excel by image name."""
    owner = state.get('excel')
    if not owner:
        return {'action': 'no_registered_excel'}
    try:
        proc = psutil.Process(owner['pid'])
        if (proc.name().lower() != 'excel.exe'
                or abs(proc.create_time() - owner['created']) > 0.01):
            return {'action': 'identity_mismatch', **owner}
        if grace:
            try:
                proc.wait(timeout=grace)
                return {'action': 'exited', **owner}
            except psutil.TimeoutExpired:
                pass
        proc.kill()
        proc.wait(timeout=5)
        return {'action': 'killed', **owner}
    except psutil.NoSuchProcess:
        return {'action': 'exited', **owner}
    except psutil.Error as exc:
        return {'action': 'cleanup_error', 'error': str(exc), **owner}


def number(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def positive(value):
    return number(value) is not None and number(value) > 0


def detail_pairs(detail):
    """Preserve repeated keys, including concatenated candidate-detail seconds."""
    for field in str(detail or '').split(';'):
        if '=' in field:
            key, value = field.split('=', 1)
            yield key.strip(), value.strip()


def temporal_key(key, event):
    key = key.lower()
    if re.search(r'(^|_)(seconds?|elapsed|remaining|deadline|budget|time)(_|$)', key):
        return True
    # These legacy budget events omit the unit in their scalar field names.
    bare = {'ordinary', 'floor', 'applied', 'limit', 'grace', 'time_cap',
            'time_floor', 'size_floor', 'cone_floor', 'flat_floor'}
    return key in bare and (event.startswith('thin_layer_') or event.startswith('hsd_budget'))


def timing_checks(scenario, state, rows):
    errors, samples, remaining, budgets = [], [], [], []
    elapsed = []
    measured_keys = {'seconds', 'wall_seconds', 'elapsed_seconds', 'step_seconds',
                     'inclusive_seconds', 'exclusive_seconds', 'logging_seconds_before_run_end'}
    macro_wall = number(state.get('com_analysis_seconds'))
    case_limit = number(state.get('actual_settings', {}).get('ADAPT_THIN_CASE_LIMIT_SECONDS'))
    for index, row in enumerate(rows):
        event = row.get('event', '')
        row_elapsed = number(row.get('elapsed_s'))
        if row_elapsed is None:
            errors.append({'row': index + 2, 'error': 'missing/nonfinite elapsed_s'})
        else:
            elapsed.append(row_elapsed)
        pairs = list(detail_pairs(row.get('detail')))
        detail = dict(pairs)
        required = set()
        if event == 'thin_layer_step' and detail.get('evaluated') == '1':
            required.add('seconds')
        elif event == 'accuracy_dll_call':
            required.update(('requested_seconds', 'packed_seconds'))
        elif event == 'thin_layer_budget':
            required.update(('ordinary', 'floor', 'applied'))
            if detail.get('solver') == 'hsd':
                required.update(('remaining_seconds', 'time_fraction', 'time_cap', 'time_floor'))
        elif event in ('thin_layer_upper_hsd', 'thin_layer_lower_hsd'):
            required.update(('remaining_seconds', 'reserve_seconds'))
            if detail.get('attempts') != '0':
                required.update(('applied', 'available_seconds', 'size_floor'))
        for key in sorted(required - detail.keys()):
            errors.append({'row': index + 2, 'event': event, 'key': key,
                           'error': 'missing required numeric time/budget field'})
        # All log columns and all detail occurrences are inspected, without a sentinel exemption.
        fields = [(k, v) for k, v in row.items() if k != 'detail'] + pairs
        for key, raw in fields:
            if not temporal_key(key, event):
                continue
            # Source audit: only this attribution and this Boolean flag carry
            # textual values among the matching production diagnostic keys.
            if key == 'remaining_source':
                if raw not in ('none', 'case_limit', 'deadline', 'nominal'):
                    errors.append({'row': index + 2, 'event': event, 'key': key,
                                   'value': raw, 'error': 'invalid remaining-source attribution'})
                continue
            if key == 'from_budget' and str(raw).lower() in ('true', 'false'):
                continue
            value = number(raw)
            if value is None:
                errors.append({'row': index + 2, 'event': event, 'key': key,
                               'value': raw, 'error': 'missing/malformed/nonfinite numeric time value'})
                continue
            sample = {'row': index + 2, 'event': event, 'elapsed_s': row_elapsed,
                      'key': key, 'value': value}
            samples.append(sample)
            if value < 0:
                errors.append({**sample, 'error': 'negative time/budget value'})
            if (key in measured_keys and macro_wall is not None
                    and value > macro_wall + CLOCK_SLACK):
                errors.append({**sample, 'error': 'measured duration exceeds macro wall time'})
        if 'remaining_seconds' in detail:
            value = number(detail['remaining_seconds'])
            if value is not None and row_elapsed is not None:
                remaining.append({'row': index + 2, 'event': event, 'elapsed_s': row_elapsed,
                                  'value': value, 'source': detail.get('remaining_source')})
                if detail.get('remaining_source') != 'case_limit':
                    errors.append({'row': index + 2, 'error': 'case limit was not used',
                                   'source': detail.get('remaining_source')})
                if case_limit is not None and value > case_limit + CLOCK_SLACK:
                    errors.append({'row': index + 2, 'error': 'remaining exceeds case limit'})
                if detail.get('remaining_source') == 'case_limit' and case_limit is not None:
                    expected_remaining = max(0.0, case_limit - row_elapsed)
                    if abs(value - expected_remaining) > CLOCK_SLACK:
                        errors.append({'row': index + 2,
                                       'error': 'remaining does not match case limit minus run elapsed',
                                       'value': value, 'expected': expected_remaining,
                                       'elapsed_s': row_elapsed, 'case_limit_seconds': case_limit})
        if event == 'thin_layer_budget':
            budget = {'row': index + 2, 'elapsed_s': row_elapsed, **detail}
            budgets.append(budget)
            # This event describes pre-clamp floors. Final applied budgets are
            # checked below; a floor can legitimately exceed remaining time.
        if event in ('thin_layer_upper_hsd', 'thin_layer_lower_hsd'):
            applied = number(detail.get('applied'))
            available = number(detail.get('available_seconds'))
            if (applied is not None and available is not None
                    and applied > available + CLOCK_SLACK):
                errors.append({'row': index + 2, 'error': 'logged HSD applied allowance exceeds available time'})
        # Native request/packed values must fit the whole case, even without a logged remaining value.
        if event == 'accuracy_dll_call' and case_limit is not None:
            for key in ('requested_seconds', 'packed_seconds'):
                value = number(detail.get(key))
                if value is not None and value > case_limit + CLOCK_SLACK:
                    errors.append({'row': index + 2, 'key': key,
                                   'error': 'native budget exceeds case limit'})
    for previous, current in zip(remaining, remaining[1:]):
        if previous['source'] != current['source']:
            continue
        delta_elapsed = current['elapsed_s'] - previous['elapsed_s']
        expected = max(0.0, previous['value'] - delta_elapsed)
        if abs(current['value'] - expected) > CLOCK_SLACK:
            errors.append({'row': current['row'], 'error': 'remaining clock discontinuity',
                           'previous': previous, 'current': current, 'expected': expected})
    before = parts(state.get('clock_hook_before', ''))
    after = parts(state.get('clock_hook_after', ''))
    before_elapsed = number(before.get('elapsed'))
    start_clock = number(state.get('clock_at_run_start'))
    end_clock = number(state.get('clock_at_run_end'))
    pre_arm = number(state.get('clock_before_arm'))
    arm_wall = number(state.get('arm_com_seconds'))
    arm_to_macro = number(state.get('arm_to_macro_entry_seconds'))
    after_elapsed = number(after.get('elapsed'))
    wrap_count = number(after.get('wraps'))
    last_elapsed = number(rows[-1].get('elapsed_s')) if rows else None
    # B includes pre-arm capture -> arming -> diagnostic origin, plus the tail
    # from final run_end to end_clock. Thus arm -> log origin is in [0, B].
    origin_bound = (end_clock - pre_arm - last_elapsed
                    if all(value is not None for value in (end_clock, pre_arm, last_elapsed)) else None)
    threshold_low = WRAP_AT - max(0.0, origin_bound) if origin_bound is not None else WRAP_AT
    threshold_high = WRAP_AT
    checks = {
        'time_fields_observed': bool(samples),
        'no_negative_nonfinite_or_discontinuous_time_values': not errors,
        'log_elapsed_all_finite_nonnegative': len(elapsed) == len(rows) and bool(rows)
            and all(value >= 0 for value in elapsed),
        'log_elapsed_monotonic': bool(elapsed) and all(b >= a for a, b in zip(elapsed, elapsed[1:])),
        'log_elapsed_fits_macro_wall': bool(elapsed) and macro_wall is not None
            and max(elapsed) <= macro_wall + CLOCK_SLACK,
        'clock_activation_continuous': all(value is not None for value in (pre_arm, start_clock, arm_wall))
            and 0 <= start_clock - pre_arm <= arm_wall + CLOCK_SLACK,
        'clock_duration_matches_macro_wall': all(value is not None for value in (start_clock, end_clock, macro_wall))
            and abs((end_clock - start_clock) - macro_wall) <= CLOCK_SLACK,
        'thin_layer_evaluated_step_seconds_observed': any(row.get('event') == 'thin_layer_step'
            and parts(row.get('detail', '')).get('evaluated') == '1'
            and number(parts(row.get('detail', '')).get('seconds')) is not None for row in rows),
        'remaining_seconds_multiple_observations': len(remaining) >= 2,
        'solver_budget_observed': bool(budgets),
        'other_failure_injection_disarmed': not any(row.get('event') in
            ('test_injection', 'test_no_iterate') for row in rows),
    }
    if scenario == 'W1':
        checks.update({
            'wrap_hook_armed_before_run': before.get('active') == '1'
                and number(before.get('wrap_at')) == WRAP_AT and number(before.get('wraps')) == 0
                and before_elapsed is not None and 0 <= before_elapsed < WRAP_AT,
            'wrap_hook_still_active_after_run': after.get('active') == '1'
                and number(after.get('wrap_at')) == WRAP_AT,
            'forced_wrap_actually_observed': wrap_count is not None and wrap_count >= 1
                and after_elapsed is not None and after_elapsed >= WRAP_AT,
            'independent_arm_to_macro_delay_bounded': arm_to_macro is not None
                and 0 <= arm_to_macro <= CLOCK_SLACK,
            'wrap_log_origin_delay_bounded': origin_bound is not None
                and -0.05 <= origin_bound <= CLOCK_SLACK,
            'analysis_rows_on_both_sides_of_wrap': bool(elapsed)
                and any(value < threshold_low - CLOCK_SLACK for value in elapsed)
                and any(value > threshold_high + CLOCK_SLACK for value in elapsed),
            'remaining_values_on_both_sides_of_wrap': any(x['elapsed_s'] < threshold_low - CLOCK_SLACK for x in remaining)
                and any(x['elapsed_s'] > threshold_high + CLOCK_SLACK for x in remaining),
            'solver_budget_after_wrap': any(number(x.get('elapsed_s')) is not None
                and x['elapsed_s'] > threshold_high + CLOCK_SLACK for x in budgets),
            'hook_raw_backward_jump_observed': number(after.get('raw_jump')) is not None
                and number(after.get('raw_jump')) < 0,
            'hook_corrected_step_nonnegative': number(after.get('corrected_step')) is not None
                and number(after.get('corrected_step')) >= 0,
        })
    else:
        checks['normal_clock_hook_disabled'] = (before.get('active') == '0'
            and after.get('active') == '0' and number(after.get('wraps')) == 0)
    return checks, {'errors': errors, 'temporal_samples': samples,
                    'remaining_samples': remaining, 'solver_budgets': budgets,
                    'clock_hook_before': before, 'clock_hook_after': after,
                    'arm_to_log_origin_upper_bound_seconds': origin_bound,
                    'wrap_threshold_interval_log_seconds': [threshold_low, threshold_high],
                    'continuity_slack_seconds': CLOCK_SLACK}


def grade(scenario, state, rows, *, timed_out, exitcode, log_error):
    ends = [(index, parts(row.get('detail', ''))) for index, row in enumerate(rows)
            if row.get('event') == 'run_end']
    terminal = ends[0][1].get('status') if len(ends) == 1 else None
    lower, upper = state.get('lower_fs'), state.get('upper_fs')
    checks = {
        'worker_finished': state.get('phase') == 'FINISHED' and exitcode == 0,
        'within_wall_limit': not timed_out,
        'no_driver_macro_or_cleanup_errors': not any(state.get(key) for key in
            ('worker_error', 'macro_error', 'cleanup_errors')),
        'settings_verified': state.get('settings_mismatch') == {},
        'run_returned': state.get('run_returned') is True,
        'analysis_log_read': log_error is None and bool(rows),
        'one_run_start': sum(row.get('event') == 'run_start' for row in rows) == 1,
        'one_final_run_end': len(ends) == 1 and ends[0][0] == len(rows) - 1,
        'certified_run_end': terminal in CERTIFIED,
        'finite_positive_ordered_fs': positive(lower) and positive(upper)
            and number(lower) <= number(upper),
    }
    timing, evidence = timing_checks(scenario, state, rows)
    checks.update(timing)
    return {**state, 'scenario': scenario, 'case': CASE['case'],
            'status': 'PASS' if all(checks.values()) else 'FAIL',
            'run_end_status': terminal, 'checks': checks,
            'failed_checks': [key for key, passed in checks.items() if not passed],
            'timing_evidence': evidence, 'log_error': log_error,
            'timed_out': timed_out, 'worker_exitcode': exitcode}


def run_scenario(source, out, scenario, limit):
    refuse_existing_excel()
    folder = out / f'scenario_{scenario}'
    folder.mkdir()
    workbook = folder / f'repro_rc9c_{scenario}.xlsm'
    dll_copy = folder / 'rpfem_native_x64.dll'
    shutil.copy2(source, workbook)
    shutil.copy2(source.with_name('rpfem_native_x64.dll'), dll_copy)
    copied_dll_hash = sha256(dll_copy)
    proc = mp.get_context('spawn').Process(target=worker, args=(workbook, folder, scenario))
    started, started_epoch = time.monotonic(), time.time()
    proc.start()
    interrupted, monitor_error, startup_owner = False, None, None
    try:
        while proc.is_alive() and time.monotonic() - started < limit:
            state = read_state(folder)
            if state.get('phase') == 'DISPATCHING' and not startup_owner:
                startup_owner = discover_startup_excel(started_epoch)
                if startup_owner:
                    write_json(folder / 'startup_excel.json', startup_owner)
            proc.join(timeout=min(0.5, max(0.0, limit - (time.monotonic() - started))))
    except KeyboardInterrupt:
        interrupted = True
    except Exception:
        interrupted, monitor_error = True, traceback.format_exc()
    timed_out = proc.is_alive() and not interrupted
    state = read_state(folder)
    if not state.get('excel'):
        startup_owner = startup_owner or discover_startup_excel(started_epoch)
        if startup_owner:
            state['excel'] = startup_owner
    cleanup = kill_owned_excel(state, grace=0 if proc.is_alive() else 3)
    if proc.is_alive():
        proc.terminate()
        proc.join(timeout=5)
        if proc.is_alive():
            proc.kill()
            proc.join(timeout=5)
    state = read_state(folder)
    if interrupted:
        state['worker_error'] = monitor_error or 'Interrupted by operator'
    if cleanup['action'] in ('identity_mismatch', 'cleanup_error', 'killed'):
        state.setdefault('cleanup_errors', []).append(str(cleanup))
    rows, log_error, log_path = [], None, None
    try:
        logs = sorted((folder / 'RPFEM_logs').glob('*_analysis*.tsv'))
        if len(logs) != 1:
            raise RuntimeError(f'Expected one analysis log, found {len(logs)}: {logs}')
        log_path = folder / 'analysis.tsv'
        shutil.copy2(logs[0], log_path)
        rows = read_log(log_path)
    except Exception:
        log_error = traceback.format_exc()
    result = grade(scenario, state, rows, timed_out=timed_out,
                   exitcode=proc.exitcode, log_error=log_error)
    result.update(workbook=str(workbook), analysis_log=str(log_path) if log_path else None,
                  excel_cleanup=cleanup, wall_seconds=time.monotonic() - started,
                  wall_limit_seconds=limit)
    result['checks']['copied_dll_unchanged'] = copied_dll_hash == sha256(dll_copy)
    result['failed_checks'] = [key for key, passed in result['checks'].items() if not passed]
    result['status'] = 'PASS' if not result['failed_checks'] else 'FAIL'
    write_json(folder / 'result.json', result)
    return result, interrupted


def compare_runs(w1, w0):
    left, right = w1.get('actual_settings', {}), w0.get('actual_settings', {})
    missing = object()
    differences = {}
    for key in sorted(left.keys() | right.keys()):
        a, b = left.get(key, missing), right.get(key, missing)
        if a is missing and key in SEISMIC_DEFAULTS:
            a = SEISMIC_DEFAULTS[key]
        if b is missing and key in SEISMIC_DEFAULTS:
            b = SEISMIC_DEFAULTS[key]
        if a is missing or b is missing or not same(a, b):
            differences[key] = {'W1': '<absent>' if a is missing else a,
                                'W0': '<absent>' if b is missing else b}
    tolerance = number(right.get('FS_TOL'))
    checks = {
        'both_scenarios_pass': w0.get('status') == 'PASS' and w1.get('status') == 'PASS',
        'all_settings_identical': bool(left) and bool(right) and not differences,
        'same_valid_fs_tol': tolerance is not None and 0 < tolerance <= 0.0001
            and same(left.get('FS_TOL'), tolerance),
        'both_certified': all(row.get('run_end_status') in CERTIFIED for row in (w0, w1)),
    }
    deltas = {}
    for key in ('lower_fs', 'upper_fs'):
        a, b = number(w1.get(key)), number(w0.get(key))
        delta = a - b if a is not None and b is not None else None
        deltas[key] = delta
        checks[key + '_equal_within_FS_TOL'] = (delta is not None and tolerance is not None
                                               and abs(delta) <= tolerance)
    return {'status': 'PASS' if all(checks.values()) else 'FAIL', 'checks': checks,
            'FS_TOL': tolerance, 'endpoint_deltas_W1_minus_W0': deltas,
            'settings_differences': differences, 'excluded_settings': []}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, required=True,
                        help='Operator-prepared rc9c .xlsm; never modifies this source')
    parser.add_argument('--out', type=Path, required=True,
                        help='New/empty subdirectory of this driver folder')
    parser.add_argument('--scenario', choices=('all', 'W0', 'W1'), default='all',
                        help='Full acceptance requires all: W0 normal then W1 120s forced wrap')
    parser.add_argument('--wall-limit', type=float, default=2400, metavar='SECONDS',
                        help='Combined startup/mesh/run/cleanup wall budget (default: 2400)')
    args = parser.parse_args(argv)
    source, out = args.workbook.resolve(), args.out.resolve()
    if not math.isfinite(args.wall_limit) or args.wall_limit <= 360:
        parser.error('--wall-limit must be finite and greater than 360 seconds')
    if not source.is_file() or source.suffix.lower() != '.xlsm':
        parser.error('--workbook must name an existing operator-prepared rc9c .xlsm')
    dll = source.with_name('rpfem_native_x64.dll')
    if not dll.is_file():
        parser.error('Place the unchanged rpfem_native_x64.dll beside the source workbook')
    if ROOT not in out.parents:
        parser.error('--out must be a subdirectory of this driver folder')
    if out.exists() and (not out.is_dir() or any(out.iterdir())):
        parser.error('--out must be new or empty; use a new directory for every invocation')
    if out in source.parents or out == source.parent:
        parser.error('The source workbook must be outside the output directory')
    refuse_existing_excel()
    scenarios = ('W0', 'W1') if args.scenario == 'all' else (args.scenario,)
    source_hashes = {'workbook': sha256(source), 'dll': sha256(dll)}
    out.mkdir(parents=True, exist_ok=True)
    summary = {
        'started_utc': utc(), 'case': CASE['case'], 'source_workbook': str(source),
        'source_sha256_before': source_hashes,
        'embedded_input_sha256': hashlib.sha256(json.dumps(CASE, ensure_ascii=False,
                                      sort_keys=True).encode('utf-8')).hexdigest(),
        'input_origin': 'rc9b driver embedded P02A_THIN; original rc9a input retained',
        'preset_macro': PRESET_MACRO, 'identical_short_overrides': SHORT,
        'clock_hook_macro': 'RPX_TestClockWrap', 'forced_wrap_at_seconds': WRAP_AT,
        'wall_limit_seconds': args.wall_limit,
        'wall_limit_scope': 'combined; 30 seconds cleanup reserve per scenario',
        'status': 'RUNNING', 'scenarios': {},
    }
    summary_path = out / 'repro_summary.json'
    write_json(summary_path, summary)
    started = time.monotonic()
    deadline = started + args.wall_limit
    for index, scenario in enumerate(scenarios):
        remaining = len(scenarios) - index
        allowance = max(0.0, (deadline - time.monotonic() - 30 * remaining) / remaining)
        print(f'Scenario {scenario}: supervised allowance {allowance:g}s', flush=True)
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
    summary['W1_vs_W0'] = (compare_runs(completed['W1'], completed['W0'])
        if {'W0', 'W1'} <= completed.keys() else
        {'status': 'INCOMPLETE', 'reason': 'Both W0 and W1 are required for acceptance'})
    try:
        summary['source_sha256_after'] = {'workbook': sha256(source), 'dll': sha256(dll)}
        summary['source_workbook_and_dll_unchanged'] = summary['source_sha256_after'] == source_hashes
    except Exception:
        summary['source_workbook_and_dll_unchanged'] = False
        summary['source_integrity_error'] = traceback.format_exc()
    failures = (len(completed) != len(scenarios)
        or any(result['status'] == 'FAIL' for result in completed.values())
        or summary['W1_vs_W0']['status'] == 'FAIL'
        or not summary['source_workbook_and_dll_unchanged'])
    summary.update(status='FAIL' if failures else
                   'INCOMPLETE' if summary['W1_vs_W0']['status'] == 'INCOMPLETE' else 'PASS',
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
