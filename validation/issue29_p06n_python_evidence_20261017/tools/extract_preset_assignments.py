"""Read only the literal AccuracySet assignments in the Fix4 preset call chains.

This does not execute Excel/VBA or infer workbook-dependent EnsureSettingRow defaults.
The only writes are the requested JSON evidence file in the scratch directory.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import subprocess
from pathlib import Path


def literal(raw: str):
    raw = raw.strip()
    if raw.startswith('"') and raw.endswith('"'):
        return raw[1:-1].replace('""', '"')
    if not re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?#?', raw):
        raise ValueError(f'Not a supported literal: {raw}')
    value = float(raw.rstrip('#'))
    return int(value) if value.is_integer() else value


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo.resolve(strict=True)
    source = repo / 'versions/G1R20_Fix4/src/vba/RPX_UI.bas'
    source_bytes = source.read_bytes()
    lines = source_bytes.decode('utf-8-sig').splitlines()
    procedures = {}
    current = None
    for number, line in enumerate(lines, 1):
        start = re.match(r'\s*(?:Public|Private) Sub (\w+)\(', line, re.I)
        if start:
            current = start.group(1)
            procedures[current] = []
        if current:
            procedures[current].append((number, line))
        if re.match(r'\s*End Sub\s*$', line, re.I):
            current = None
    presets = ('RPX_AccuracyPreset', 'RPX_Speed2pctPreset', 'RPX_G20ImprovementPreset')
    action_re = re.compile(
        r'\bCall\s+(RPX_AccuracyPreset|RPX_Speed2pctPreset)\b'
        r'|\bAccuracySet\s+"([^"]+)"\s*,\s*("(?:[^"]|"")*"|[^:\r\n]+)', re.I)

    def collect(name, stack=()):
        if name in stack:
            raise ValueError('Recursive preset call chain')
        actions = []
        for number, raw in procedures[name]:
            # These inspected procedures have no apostrophes inside a string.
            line = raw.split("'", 1)[0]
            for match in action_re.finditer(line):
                if match.group(1):
                    actions.extend(collect(match.group(1), stack + (name,)))
                else:
                    actions.append({'key': match.group(2), 'value': literal(match.group(3)),
                                    'procedure': name, 'line': number})
        return actions

    profiles = {}
    for preset in presets:
        actions = collect(preset)
        final = {action['key']: action for action in actions}
        profiles[preset] = {'final_assignments': {key: item['value'] for key, item in sorted(final.items())},
                            'last_assignment_source': {key: {'procedure': item['procedure'], 'line': item['line']}
                                                       for key, item in sorted(final.items())},
                            'ordered_assignments': actions}
    speed = profiles['RPX_Speed2pctPreset']['final_assignments']
    improvement = profiles['RPX_G20ImprovementPreset']['final_assignments']
    comparison = [{'key': key, 'speed': speed.get(key), 'improvement': improvement.get(key),
                   'different': speed.get(key) != improvement.get(key)}
                  for key in sorted(speed.keys() | improvement.keys())]
    revision = subprocess.run(['git', 'rev-parse', 'HEAD'], cwd=repo, text=True,
                              capture_output=True, check=True).stdout.strip()
    reference_paths = [
        'versions/G1R20_Fix4/src/vba/' + name for name in
        ('RPX_UI.bas', 'RPX_Adapt.bas', 'RPX_AccuracyNative.bas', 'RPX_Fs.bas', 'RPX_Robust.bas')
    ] + [
        'validation/G1R19_python26_20261016/python/accuracy_driver.py',
        'validation/G1R19_python26_20261016/python/experiment_core.py',
        'validation/G1R19_python26_dual_linear_repair_20261016/G1R20_RECIPE.md',
        'versions/G1R20/README_JA.md', 'versions/G1R20/TEST_RESULTS.md'
    ]
    reference_sources = []
    for path in reference_paths:
        raw = (repo / path).read_bytes()
        reference_sources.append({'path': path, 'bytes': len(raw),
                                  'sha256': hashlib.sha256(raw).hexdigest()})
    code_bytes = Path(__file__).read_bytes()
    result = {'schema': 'RPFEM_LITERAL_PRESET_ASSIGNMENTS_V1',
              'created_utc': dt.datetime.now(dt.timezone.utc).isoformat(),
              'repo': str(repo), 'git_head': revision,
              'source': {'path': source.relative_to(repo).as_posix(), 'bytes': len(source_bytes),
                         'sha256': hashlib.sha256(source_bytes).hexdigest()},
              'reference_sources': reference_sources,
              'extractor': {'path': str(Path(__file__).resolve()), 'bytes': len(code_bytes),
                            'sha256': hashlib.sha256(code_bytes).hexdigest()},
              'scope': 'Literal AccuracySet assignments and explicit preset calls only; no Excel, no solve.',
              'limitations': [
                  'EnsureSettingRow/EnsureG1R18Settings only create missing settings; their conditional effects are not executed.',
                  'Existing workbook settings not explicitly assigned by these presets may remain in effect.',
                  'This table does not prove the settings actually used by an Issue #29 Excel session.'
              ],
              'profiles': profiles, 'comparison': comparison}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'output': str(args.output.resolve()), 'source': result['source'],
                      'profiles': {k: len(v['final_assignments']) for k, v in profiles.items()},
                      'differences': [item for item in comparison if item['different']]}, ensure_ascii=False))


if __name__ == '__main__':
    main()
