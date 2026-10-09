"""Read XLSM XML only and compare saved tables to their native witnesses.

Does not use a spreadsheet writer, replace VBA, or attach to a running Excel.
XML comparisons allow only Excel numeric serialization error. Native solver's
exact owner/model guards remain the identity proof.
"""
from pathlib import Path
import zipfile, xml.etree.ElementTree as ET, posixpath, json
import numpy as np
from prepare_inputs import ROOT
from audit_batch import read_state
from case_locations import case_folder,evidence_root

NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def sheets(path):
    with zipfile.ZipFile(path) as z:
        rel={r.attrib['Id']:r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            shared=[''.join(t.text or '' for t in e.iter('{'+NS['s']+'}t')) for e in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        out={}
        for e in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets',NS):
            name=e.attrib['name'];target=rel[e.attrib['{'+RNS+'}id']]
            member=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
            cells={}
            for c in ET.fromstring(z.read(member)).findall('.//s:sheetData/s:row/s:c',NS):
                v=c.find('s:v',NS);kind=c.attrib.get('t')
                if kind=='inlineStr':value=''.join(t.text or '' for t in c.iter('{'+NS['s']+'}t'))
                elif v is None:value=None
                elif kind=='s':value=shared[int(v.text)]
                elif kind in ['str','e']:value=v.text
                else:value=float(v.text)
                cells[c.attrib['r']]=value
            out[name]=cells
        return out

def column(n):
    out=''
    while n:n,k=divmod(n-1,26);out=chr(65+k)+out
    return out

def matrix(cells,col,row,shape):
    return np.array([[cells[column(col+j)+str(row+i)] for j in range(shape[1])] for i in range(shape[0])],dtype=float)

def check(case):
    folder=case_folder(case);r=json.loads((folder/'status.json').read_text(encoding='utf-8'))
    assert r['status'] in ['PASS_TARGET','FS_AVAILABLE_PARTIAL','AUDIT_FAILED'],r['status']
    rec=json.loads((ROOT/'inputs'/f'{case}.json').read_text(encoding='utf-8'))
    tabs=sheets(evidence_root(case)/r['workbook']);checks={}
    for col,table in [(1,'points'),(5,'regions'),(10,'faces')]:
        for i,values in enumerate(rec[table]):
            for j,v in enumerate(values):
                actual=tabs['要素定義'][column(col+j)+str(i+7)]
                assert actual==v,(case,table,i,j,actual,v)
    assert np.array_equal(matrix(tabs['材料データ'],1,5,(len(rec['materials']),5)),np.array(rec['materials']))
    checks['input_tables_readback']=True
    for i,side in enumerate(['upper','lower']):
        idx=1 if side=='upper' else 0
        m,x,meta=read_state(folder/f'witness_{side}.bin',rec['source_case']['policy'],r['witness_phase'][idx])
        nc,ec,vc,sc=(1,1,8,None) if side=='upper' else (6,27,None,34)
        nodes=matrix(tabs['節点データ'],nc,2,(len(m.xy),3))
        elems=matrix(tabs['要素データ'],ec,2,(len(m.tri),6))
        assert np.array_equal(nodes[:,0],np.arange(1,len(m.xy)+1))
        assert np.allclose(nodes[:,1:],m.xy,rtol=3e-14,atol=1e-12)
        expected=np.column_stack([np.arange(1,len(m.tri)+1),m.tri+1,m.material_id+1,m.rigid_id+1])
        assert np.array_equal(elems,expected),(case,side,'topology/material output mismatch')
        if side=='upper':
            actual=matrix(tabs['要素データ'],vc,2,(len(m.tri),12));wanted=x[:12*len(m.tri)].reshape(-1,12)
        else:
            actual=matrix(tabs['要素データ'],sc,2,(len(m.tri),18));wanted=x[:18*len(m.tri)].reshape(-1,18)*m.stress_scale
        assert np.allclose(actual,wanted,rtol=3e-14,atol=1e-12),(case,side,'saved field mismatch')
        checks[side]={'pass':True,'nodes':len(m.xy),'elements':len(m.tri),'max_saved_field_delta':float(np.max(abs(actual-wanted)))}
    assert np.allclose([tabs['解析結果']['B4'],tabs['解析結果']['C4']],[r['lower_fs'],r['upper_fs']],rtol=3e-14,atol=0.)
    out={'case':case,'pass':True,'checks':checks,'scope':'saved XML tables vs native witnesses; exact native owner guards plus bounded Excel numeric serialization comparison'}
    (folder/'saved_output_check.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
    return out

if __name__=='__main__':
    done=[]
    for inp in sorted((ROOT/'inputs').glob('*.json')):
        p=case_folder(inp.stem)/'status.json'
        if not p.exists():continue
        r=json.loads(p.read_text(encoding='utf-8'))
        if r['status'] in ['PASS_TARGET','FS_AVAILABLE_PARTIAL','AUDIT_FAILED']:done.append(check(r['case']))
    print(json.dumps(done,ensure_ascii=False))
