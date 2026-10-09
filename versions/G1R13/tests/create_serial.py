from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
s=(ROOT/'tests/run_full.py').read_text(encoding='utf-8')
s=s.replace("folder=ROOT/'cases'/case","folder=ROOT/'native_retry'/case")
s=s.replace('folder.mkdir(exist_ok=True)','folder.mkdir(parents=True,exist_ok=True)')
s=s.replace("write(ROOT/'results'/f'{case}.json',out)","write(ROOT/'results'/f'retry_{case}.json',out)")
s=s[:s.index("if __name__=='__main__':")]+"""if __name__=='__main__':
    worker('D07A')
    worker('D03N')
"""
(ROOT/'tests/run_serial.py').write_text(s,encoding='utf-8')
