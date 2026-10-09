"""Compare the production P06N coarse geometry to the catalog adapter.

Uses separate diagnostic workbook copies, never attaches to batch Excel processes.
No FEM solve and no change to production source or in-flight case inputs.
"""
from pathlib import Path
import json, shutil, struct, hashlib, sys
import numpy as np
from scipy.spatial import cKDTree
import win32com.client
from prepare_inputs import ROOT, SOURCE, SOURCE_SHA, REPO
sys.path.insert(0, str(REPO/'versions/G1R12/tests'))
from native_helpers import compile_book

def read(path):
    data=path.read_bytes();assert data[:8]==b'RP25GEOM'
    nn,ne=struct.unpack_from('<ii',data,8)
    xy=np.frombuffer(data,dtype='<f8',count=2*nn,offset=16).copy().reshape(nn,2)
    tr=np.frombuffer(data,dtype='<i4',count=4*ne,offset=16+16*nn).copy().reshape(ne,4)
    assert len(data)==16+16*nn+16*ne
    return xy,tr

def run(target):
    folder=ROOT/('mesh_identity_target'+str(target));folder.mkdir(exist_ok=True)
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
    book=None
    try:
        for tag in ['production','catalog_adapter']:
            path=folder/(tag+'.xlsm');assert not path.exists();shutil.copy2(SOURCE,path)
            book=app.Workbooks.Open(str(path),UpdateLinks=0)
            cm=book.VBProject.VBComponents.Add(1);cm.Name='MeshIdentity'
            cm.CodeModule.AddFromString('\r\n'.join((ROOT/'tests/MeshIdentity.bas').read_text(encoding='utf-8').splitlines()[1:]))
            compile_book(app,book)
            if tag=='catalog_adapter':
                rec=json.loads((ROOT/'inputs/P06N.json').read_text(encoding='utf-8'))
                ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
                for first,last,name in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
                    data=tuple(tuple(v) for v in rec[name]);rng=ws.Range(first+'7:'+last+str(6+len(data)))
                    rng.Value2=data;assert rng.Value2==data
            value=app.Run("'"+book.Name+"'!MeshIdentityExport",str(folder/(tag+'.bin')),target)
            assert value.startswith('OK '),value
            book.Close(SaveChanges=False);book=None
        a,at=read(folder/'production.bin');b,bt=read(folder/'catalog_adapter.bin')
        dist,indices=cKDTree(a).query(b);tolerance=1e-11
        bijective=len(set(indices))==len(a)==len(b) and float(dist.max())<=tolerance
        def canon(tr,index):
            return set(tuple(sorted(index[row[:3]]))+tuple(row[3:]) for row in tr)
        if bijective:
            ac=canon(at,np.arange(len(a)));bc=canon(bt,indices)
            common=len(ac&bc);same=ac==bc
        else:common=None;same=False
        out={'scope':'mesh-only diagnostic; no FEM solve','target':target,
             'production_nodes':len(a),'adapter_nodes':len(b),'production_elements':len(at),'adapter_elements':len(bt),
             'coordinate_matching_tolerance':tolerance,'max_nearest_distance':float(dist.max()),
             'nodes_bijective_within_tolerance':bijective,'same_triangles_and_material_ids':same,
             'common_triangles_with_material_ids':common,'production_source_unchanged':hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA,
             'interpretation':'Different numbering/segment order may change triangulation; no conclusion about solver accuracy from this mesh-only check.'}
        (folder/'comparison.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(out,ensure_ascii=False))
    finally:
        if book is not None:book.Close(SaveChanges=False)
        app.Quit()

if __name__=='__main__':run(int(sys.argv[1]) if len(sys.argv)>1 else 538)
