"""Create independently extractable workbook ZIPs below the GitHub blob API limit."""
import hashlib, json, zipfile
from pathlib import Path
from prepare_inputs import ROOT

def sha(raw): return hashlib.sha256(raw).hexdigest()

def run():
    folder=ROOT/'delivery'
    verified=json.loads((folder/'delivery_verification.json').read_text(encoding='utf-8'))
    assert verified['pass'] and verified['native_workbooks']==25
    original=folder/'RPFEM_G1R12_Native25_Workbooks_20261009.zip'
    packages=json.loads((folder/'package_summary.json').read_text(encoding='utf-8'))
    expected=next(x for x in packages if x['name']==original.name)
    assert sha(original.read_bytes())==expected['sha256']
    results=[];covered=[]
    with zipfile.ZipFile(original) as source:
        books=json.loads(source.read('WORKBOOK_MANIFEST.json'));assert len(books)==25
        for number,group in enumerate([books[:12],books[12:]],1):
            name=f'RPFEM_G1R12_Native25_Workbooks_{number:02d}_20261009.zip'
            target=folder/name;assert not target.exists()
            payload={k:source.read(k) for k in ['FINAL_REPORT.md','FINAL_RESULTS.csv']}
            for book in group:
                path=book['path'].replace('\\','/');raw=source.read(path)
                assert sha(raw)==book['sha256'];payload[path]=raw;covered.append(book['case'])
            payload['WORKBOOK_MANIFEST.json']=json.dumps(group,ensure_ascii=False,indent=2).encode('utf-8')
            cases='、'.join(x['case'] for x in group)
            note=f'''# 結果XLSM ZIP {number}/2\n\nこのZIPは{len(group)}ケース（{cases}）を含みます。2つのZIPは独立して通常どおり展開でき、結合は不要です。両方を取得すると25ケースすべてが揃います。\n\nFINAL_REPORT.md・FINAL_RESULTS.csvは全25ケースの集計です。WORKBOOK_MANIFEST.jsonがこのZIPに含まれるブックの一覧です。FAILEDのブックは最終Fs未取得、FS_AVAILABLE_PARTIALは監査済みFsを取得したものの精度目標が未達です。\n\n各XLSMの内容は検証済みの一括ZIPとバイト単位で同一です。通常のExcel解析にPythonやVBProjectアクセスは不要です。元の一括ZIPはローカル成果物として保全し、GitHub APIのサイズ制限に対応してこの2つのZIPを公開しています。\n'''
            payload['PART_README_JA.md']=note.encode('utf-8')
            manifest=[{'path':k,'bytes':len(v),'sha256':sha(v)} for k,v in sorted(payload.items())]
            payload['FILE_MANIFEST.json']=json.dumps(manifest,ensure_ascii=False,indent=2).encode('utf-8')
            with zipfile.ZipFile(target,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=9) as z:
                for k,v in sorted(payload.items()):z.writestr(k,v)
            with zipfile.ZipFile(target) as z:
                assert z.testzip() is None
                assert set(z.namelist())==set(payload)
                for k,v in payload.items():assert z.read(k)==v
            assert target.stat().st_size<29_000_000
            results.append({'name':name,'bytes':target.stat().st_size,'sha256':sha(target.read_bytes()),'cases':[b['case'] for b in group]})
    assert len(covered)==len(set(covered))==25
    assert set(covered)=={b['case'] for b in books}
    evidence=next(x for x in packages if 'Evidence' in x['name'])
    public=[evidence,*results]
    (folder/'public_package_summary.json').write_text(json.dumps(public,ensure_ascii=False,indent=2),encoding='utf-8')
    result={'pass':True,'all25_cases_covered_once':True,'native_workbooks_byte_identical_to_original_zip':True,'independent_archives_no_join_required':True,'original_full_zip':expected,'public_packages':public}
    (folder/'public_package_verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    with (folder/'SHA256SUMS.txt').open('a',encoding='utf-8') as f:
        for p in results:f.write(p['sha256']+'  '+p['name']+'\n')
    print(json.dumps(result,ensure_ascii=False))

if __name__=='__main__':run()
