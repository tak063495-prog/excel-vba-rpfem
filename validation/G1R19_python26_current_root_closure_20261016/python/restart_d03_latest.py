"""Preserve the owned D03 run, then launch the current audited upper continuation."""
from pathlib import Path
import hashlib, json, subprocess, sys, time
from datetime import datetime, timezone
import psutil

repo = Path(__file__).resolve().parents[3]
results = repo / 'validation' / 'accuracy_20261013' / 'results'
old = results / 'd_combined'
seed = old / 'D03N'
output = results / 'd03_latest'
if output.exists():
    raise RuntimeError(f'Restart output already exists: {output}')
records = json.loads((old / 'd_suite_progress.json').read_text(encoding='utf-8'))
record = next(r for r in records if r['case'] == 'D03N')
files = [seed / f'best_{side}_{kind}.npz' for side in ('lower', 'upper') for kind in ('field', 'mesh')]
before = {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
owned = psutil.Process(record['pid'])
targets = [p for p in owned.children(recursive=True) if p.name().lower() == 'python.exe']
for proc in targets or [owned]:
    cmd = proc.cmdline()
    if '--case' not in cmd or cmd[cmd.index('--case') + 1] != 'D03N':
        raise RuntimeError(f'Unowned case process {proc.pid}')
    if '--output' not in cmd or Path(cmd[cmd.index('--output') + 1]).resolve() != old.resolve():
        raise RuntimeError(f'Unowned output process {proc.pid}')
    proc.terminate()
    proc.wait(timeout=30)
after = {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
if before != after:
    raise RuntimeError('Saved independent witnesses changed during interruption')
stamp = time.time_ns()
audit = {'reason': 'Restart latest full-original Farkas guards and equivalent CSR cone normalization; retain independent lower witness',
         'utc': datetime.now(timezone.utc).isoformat(), 'old_command': record['command'],
         'witnesses_sha256': before, 'all_witnesses_unchanged': True,
         'history_retained_in_place': str(old)}
(old / f'D03N_interruption_latest_{stamp}.json').write_text(json.dumps(audit, indent=2), encoding='utf-8')
command = [sys.executable, '-X', 'utf8', str(Path(__file__).with_name('accuracy_driver.py')),
           '--case', 'D03N', '--output', str(output), '--seed', str(seed), '--recheck',
           '--upper-only', '--cycles', '8', '--max-elements', '12000']
print(json.dumps({'command': command, 'interruption': audit}, ensure_ascii=False), flush=True)
subprocess.run(command, cwd=repo, check=True)
