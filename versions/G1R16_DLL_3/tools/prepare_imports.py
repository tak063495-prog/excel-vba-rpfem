from prepare import ROOT,sha,write
def main():
    changed=[]
    for p in sorted((ROOT/'src/vba').glob('*.bas')):
        baseline=ROOT/'baseline_sources'/p.name
        if baseline.exists() and p.read_bytes()==baseline.read_bytes():continue
        assert p.name not in ['ThisWorkbook.bas'] and not p.name.startswith('Sheet')
        text=p.read_text(encoding='utf-8');data=text.replace('\n','\r\n').encode('cp932',errors='strict')
        assert data.decode('cp932').replace('\r\n','\n')==text
        assert max(map(len,text.splitlines()),default=0)<1024,p.name
        target=ROOT/'import'/p.name;target.write_bytes(data);changed.append(p.stem)
    write(ROOT/'results/import_modules.json',{'modules':changed,'encoding':'CP932 strict round trip','line_endings':'CRLF'})
    print('IMPORTS',','.join(changed),flush=True)
if __name__=='__main__':main()
