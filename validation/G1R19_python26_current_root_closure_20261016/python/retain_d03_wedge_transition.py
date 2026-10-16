"""Stop only the owned D03 continuation; preserve witnesses and all history."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib
import json
import time
import psutil

repo = Path(__file__).resolve().parents[3]
folder = repo / 'validation/accuracy_20261013/results/d03_latest'
case = folder / 'D03N'
owned = psutil.Process(18816)
command = owned.cmdline()
assert owned.name().lower() == 'python.exe'
assert command[command.index('--case') + 1] == 'D03N'
assert Path(command[command.index('--output') + 1]).resolve() == folder.resolve()
parents = []
for parent in owned.parents():
    parent_command = parent.cmdline()
    if parent.name().lower() != 'python.exe':
        break
    if any(Path(token).name == 'restart_d03_latest.py' for token in parent_command):
        parents.append(parent)
    elif '--case' in parent_command and parent_command[parent_command.index('--case') + 1] == 'D03N' and '--output' in parent_command and Path(parent_command[parent_command.index('--output') + 1]).resolve() == folder.resolve():
        parents.append(parent)
    else:
        break
files = [case / f'best_{side}_{kind}.npz' for side in ('lower', 'upper') for kind in ('field', 'mesh')]
before = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
record = {'utc': datetime.now(timezone.utc).isoformat(), 'reason':
    'Transition to independently audited same-physical-model c0 wedge upper with complete own q8 root; retained continuation results are not promoted',
    'prior_command': command, 'pids': [owned.pid] + [p.pid for p in parents],
    'saved_witnesses_sha256': before, 'history_retained_in_place': str(folder),
    'replacement_evidence': str(repo / 'validation/accuracy_20261013/results/c0_slope_wedge_1791899277745792400/D03N')}
owned.terminate()
owned.wait(timeout=30)
for parent in parents:
    try:
        parent.terminate()
        parent.wait(timeout=30)
    except psutil.NoSuchProcess:
        pass
after = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
assert before == after, 'Saved independent fields changed during interruption'
record['all_saved_witnesses_unchanged'] = True
path = folder / f'D03N_interruption_wedge_{time.time_ns()}.json'
path.write_text(json.dumps(record, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(path)
print(json.dumps(record, ensure_ascii=False))
