"""Run the dry target-unmet cases from the published 26-case table."""
import argparse
import csv
import json
from pathlib import Path
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from accuracy_driver import REPO

parser = argparse.ArgumentParser()
parser.add_argument('--output', type=Path, required=True)
parser.add_argument('--cycles', type=int, default=8)
parser.add_argument('--max-elements', type=int, default=12000)
parser.add_argument('--cases', nargs='*')
parser.add_argument('--seed-root',type=Path)
parser.add_argument('--skip-seed-recheck',action='store_true',
                    help='Continue stored witnesses; the driver still validates every retained root identity')
parser.add_argument('--fan-lower',action='store_true')
parser.add_argument('--parallel',type=int,default=1)
args = parser.parse_args()
table = REPO / 'docs/benchmarks/G1R17_native26_20261013/data/RESULTS.csv'
rows = list(csv.DictReader(table.open(encoding='utf-8-sig', newline='')))
cases = args.cases or [r['case'] for r in rows if r['target_met'] == 'False' and r['case'] not in ['P06A', 'P06N']]
args.output.mkdir(parents=True, exist_ok=True)
results = []
def run_case(case):
    command = [sys.executable, '-X', 'utf8', str(Path(__file__).with_name('accuracy_driver.py')),
               '--case', case, '--output', str(args.output.resolve()), '--cycles', str(args.cycles),
               '--max-elements', str(args.max_elements)]
    if args.seed_root and (args.seed_root/case/'best_lower_field.npz').exists() and (args.seed_root/case/'best_upper_field.npz').exists():
        command += ['--seed',str((args.seed_root/case).resolve())]
        if not args.skip_seed_recheck:command.append('--recheck')
    if args.fan_lower:
        command.append('--fan-lower')
    with (args.output / (case + '.log')).open('w', encoding='utf8') as log:
        status = subprocess.run(command, stdout=log, stderr=subprocess.STDOUT).returncode
    summary = args.output / case / 'summary.json'
    result = json.loads(summary.read_text(encoding='utf8')) if summary.exists() else {}
    return {'case': case, 'returncode': status, 'target_met': result.get('target_met', False),
                    'width': result.get('relative_width_symmetric'), 'summary': str(summary),'command':command}

with ThreadPoolExecutor(max_workers=max(1,args.parallel)) as pool:
  for future in as_completed({pool.submit(run_case,case):case for case in cases}):
    results.append(future.result())
    (args.output / 'suite_progress.json').write_text(json.dumps(results, indent=2) + '\n', encoding='utf8')
    print('SUITE', results[-1], flush=True)
