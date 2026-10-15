"""Select the authoritative 62 G1R19 source modules and verify import parity."""
from pathlib import Path
import hashlib
RELEASE=Path(__file__).resolve().parents[1]
def module_paths(source_root=None):
    root=Path(source_root).resolve() if source_root else RELEASE
    paths={p.stem:p for p in sorted((root/'src/vba').glob('*.bas'))}
    if len(paths)!=62:raise ValueError(f'Expected 62 source modules: {root}')
    for name,path in paths.items():
        source=path.read_text(encoding='utf8')
        if not source.startswith(f'Attribute VB_Name = "{name}"'):raise ValueError(path)
        if (root/'import'/path.name).read_text(encoding='cp932')!=source:raise ValueError('Import mismatch: '+name)
    return paths
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
