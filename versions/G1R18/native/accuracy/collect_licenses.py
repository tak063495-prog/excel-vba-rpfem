"""Index pinned registry dependencies and copy their available license notices."""
from pathlib import Path
import argparse
import hashlib
import json
import shutil
import tomllib
from urllib.error import HTTPError, URLError
from urllib.request import urlopen


def run(cargo_home, fetch_missing=False):
    root = Path(__file__).resolve().parent
    lock = tomllib.loads((root / 'Cargo.lock').read_text(encoding='utf8'))
    cache = cargo_home / 'registry' / 'src'
    index = []
    for package in lock['package']:
        if not package.get('source', '').startswith('registry+'):
            continue
        name, version = package['name'], package['version']
        found = list(cache.glob(f'*/{name}-{version}'))
        if len(found) != 1:
            raise ValueError(f'Expected one pinned crate source: {name} {version}')
        crate = found[0]
        manifest = tomllib.loads((crate / 'Cargo.toml').read_text(encoding='utf8'))['package']
        paths = {p for pattern in ['LICENSE*', 'COPYING*', 'NOTICE*', 'COPYRIGHT*']
                 for p in crate.glob(pattern) if p.is_file()}
        declared = manifest.get('license-file')
        if declared:
            path = (crate / declared).resolve()
            if path.is_relative_to(crate.resolve()) and path.is_file():
                paths.add(path)
        output = root / 'third_party' / f'{name}-{version}'
        copied = {}
        for path in sorted(paths):
            output.mkdir(parents=True, exist_ok=True)
            target = output / path.name
            shutil.copyfile(path, target)
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            if hashlib.sha256(target.read_bytes()).hexdigest() != digest:
                raise ValueError('License copy differs from its pinned source')
            copied[path.name] = digest
        # faer package archives omit the root repository LICENSE; the pinned
        # upstream notice is obtained separately and kept in this directory.
        if not copied and output.is_dir():
            copied = {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
                      for p in output.glob('LICENSE*') if p.is_file()}
        if not copied and fetch_missing:
            repository = manifest.get('repository', '').rstrip('/').removesuffix('.git')
            vcs = json.loads((crate / '.cargo_vcs_info.json').read_text(encoding='utf8'))
            commit = vcs['git']['sha1']
            if not repository.startswith('https://github.com/') or len(commit) != 40:
                raise ValueError('A pinned public GitHub license source is required')
            relative = repository.removeprefix('https://github.com/')
            for filename in ['LICENSE', 'LICENSE-MIT', 'LICENSE.md', 'LICENSE.txt']:
                url = f'https://raw.githubusercontent.com/{relative}/{commit}/{filename}'
                try:
                    with urlopen(url, timeout=20) as response:
                        notice = response.read()
                except (HTTPError, URLError):
                    continue
                output.mkdir(parents=True, exist_ok=True)
                (output / filename).write_bytes(notice)
                digest = hashlib.sha256(notice).hexdigest()
                copied[filename] = digest
                (output / 'upstream_notice_source.json').write_text(json.dumps({
                    'url': url, 'pinned_source_commit': commit,
                    'sha256': digest, 'source': 'upstream license data only'}, indent=2), encoding='utf8')
                break
            if not copied:
                # npyz publishes its copyright and MIT declaration in README,
                # rather than a standalone LICENSE. Preserve that exact pinned
                # source and reproduce the MIT permission with its two notices.
                if name != 'npyz' or version != '0.8.4' or manifest.get('license') != 'MIT':
                    raise ValueError(f'No upstream pinned license notice found: {name} {version}')
                readme = (crate / 'README.md').read_bytes()
                required = ['Copyright 2021 Michael Lamparski', 'Copyright 2018 Pavel Potoček',
                            'provided under the terms of the MIT License']
                if not all(text in readme.decode('utf8') for text in required):
                    raise ValueError('Pinned npyz copyright declarations differ')
                output.mkdir(parents=True, exist_ok=True)
                (output / 'LICENSE_SOURCE_README.md').write_bytes(readme)
                permission = '''MIT License

Copyright 2021 Michael Lamparski
Copyright 2018 Pavel Potoček

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
'''
                (output / 'LICENSE-MIT').write_text(permission, encoding='utf8')
                copied = {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
                          for p in output.glob('LICENSE*') if p.is_file()}
                (output / 'upstream_notice_source.json').write_text(json.dumps({
                    'url': f'https://github.com/{relative}/blob/{commit}/README.md',
                    'pinned_source_commit': commit,
                    'source': 'Pinned crate README copyright/MIT declarations plus standard MIT permission',
                    'readme_sha256': hashlib.sha256(readme).hexdigest(),
                    'lockfile_checksum': package['checksum']}, indent=2), encoding='utf8')
        index.append({'name': name, 'version': version, 'spdx': manifest.get('license'),
            'repository': manifest.get('repository'), 'lockfile_checksum': package.get('checksum'),
            'available_notice_sha256': copied})
    destination = root / 'third_party' / 'locked_dependency_licenses.json'
    destination.write_text(json.dumps({'scope': 'All Cargo.lock registry packages, including optional/platform dependencies',
        'cargo_lock_sha256': hashlib.sha256((root / 'Cargo.lock').read_bytes()).hexdigest(),
        'packages': index}, indent=2), encoding='utf8')
    lines = ['# Locked dependency notices', '',
        'This index includes optional and platform-specific lockfile entries as well as Windows build dependencies.', '',
        '| Crate | Version | Declared license | Notice files |', '|---|---|---|---|']
    for p in index:
        lines.append(f"| {p['name']} | {p['version']} | {p['spdx']} | {', '.join(p['available_notice_sha256']) or 'archive has no separate notice'} |")
    (root / 'third_party' / 'NOTICES.md').write_text('\n'.join(lines)+'\n', encoding='utf8')
    print('PINNED_DEPENDENCY_LICENSE_INDEX', len(index), destination)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cargo-home', type=Path, default=Path(r'D:\Work\rpfem-rust-bootstrap\cargo'))
    parser.add_argument('--fetch-missing-license', action='store_true')
    args = parser.parse_args()
    run(args.cargo_home, args.fetch_missing_license)
