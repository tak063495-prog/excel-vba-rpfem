"""Run the dry target-unmet cases from the published 26-case table."""
import argparse
import csv
import json
from pathlib import Path
import subprocess
import sys
from accuracy_driver import REPO

parser = argparse.ArgumentParser()
parser.add_argument('--output', type=Path, required=True)
parser.add_argument('--cycles', type=int, default=8)
parser.add_argument('--max-elements', type=int, default=12000)
parser.add_argument('--cases', nargs='*')
args = parser.parse_args()
table = REPO / 'docs/benchmarks/G1R17_native26_20261013/data/RESULTS.csv'
rows = list(csv.DictReader(table.open(encoding='utf-8-sig', newline='')))
cases = args.cases or [r['case'] for r in rows if r['target_met'] == 'False' and r['case'] not in ['P06A', 'P06N']]
args.output.mkdir(parents=True, exist_ok=True)
results = []
for case in cases:
    command = [sys.executable, '-X', 'utf8', str(Path(__file__).with_name('accuracy_driver.py')),
               '--case', case, '--output', str(args.output.resolve()), '--cycles', str(args.cycles),
               '--max-elements', str(args.max_elements)]
    with (args.output / (case + '.log')).open('w', encoding='utf8') as log:
        status = subprocess.run(command, stdout=log, stderr=subprocess.STDOUT).returncode
    summary = args.output / case / 'summary.json'
    result = json.loads(summary.read_text(encoding='utf8')) if summary.exists() else {}
    results.append({'case': case, 'returncode': status, 'target_met': result.get('target_met', False),
                    'width': result.get('relative_width_symmetric'), 'summary': str(summary)})
    (args.output / 'suite_progress.json').write_text(json.dumps(results, indent=2) + '\n', encoding='utf8')
    print('SUITE', results[-1], flush=True)
