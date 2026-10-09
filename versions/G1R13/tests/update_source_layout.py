from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
for p in list((ROOT/'tests').glob('*.py'))+list(ROOT.glob('*.md')):
    if p.name=='update_source_layout.py':continue
    s=p.read_text(encoding='utf-8')
    s=s.replace("ROOT/'src/","ROOT/'src/vba/").replace("ROOT/'src'","ROOT/'src/vba'")
    if p.suffix=='.md':s=s.replace('`src/`','`src/vba/`')
    if p.name=='report_and_package.py':s=s.replace('`src/`','`src/vba/`')
    p.write_text(s,encoding='utf-8',newline='\n')
print('SOURCE LAYOUT src/vba')
