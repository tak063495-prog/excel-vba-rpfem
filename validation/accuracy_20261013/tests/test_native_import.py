"""Native binary evidence cannot bypass current model and physical audits."""
import struct
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import import_native as native
from experiment_core import np


def snapshot(q=1,factor=1.,fs=1.):
    xy=np.array([[0.,0.],[1.,0.],[0.,1.]])
    cells=np.array([[0,1,2,0,-1]],dtype='<i4')
    materials=np.array([[1.,30.,30.,18.]])
    body=np.zeros((1,2,2))
    field=np.zeros(19)
    boundary={(0,1):{'kind':'fixed','rigid':-1,'loads':np.zeros((2,2))}}
    raw=b'RP25M001'+struct.pack('<5i',3,1,1,0,q)+struct.pack('<5d',1.,factor,fs,0.,0.)
    raw+=xy.astype('<f8').tobytes()+cells.tobytes()+materials.astype('<f8').tobytes()+body.astype('<f8').tobytes()
    raw+=struct.pack('<i4i',1,0,1,1,-1)+np.zeros(4,dtype='<f8').tobytes()+struct.pack('<i',19)+field.astype('<f8').tobytes()
    mesh=SimpleNamespace(tri=cells[:,:3],materials=materials,body=body,boundary=boundary,stress_scale=1.)
    record={'witness_phase':['reference_total','reference_total'],'witness_q':[q,q],'lower_fs':fs,'upper_fs':fs}
    return raw,mesh,record


class NativeImportTests(unittest.TestCase):
    def test_binary_rejects_truncation_phase_rigid_q_and_trailing_data(self):
        raw,_,_=snapshot()
        bad_rigid=bytearray(raw);struct.pack_into('<i',bad_rigid,20,1)
        for bad in [raw[:7],raw[:-1],raw+b'x',bytes(bad_rigid),snapshot(q=0)[0],snapshot(factor=2.)[0]]:
            with self.subTest(size=len(bad)),self.assertRaises(ValueError):
                native.read_state(bad)

    def test_source_q_is_preserved_but_audit_uses_current_q8(self):
        raw,mesh,record=snapshot(q=1)
        result={'physical_pass':True,'witness':True,'value':1.000003}
        with patch.object(native.core,'mesh_from_arrays',return_value=mesh),patch.object(native,'Prepared') as program,patch.object(native,'evaluate',return_value=result):
            cert,report=native.import_side(raw,record,{},'lower')
        self.assertEqual(program.call_args.args[-1],8)
        self.assertEqual(report['source']['q'],1)
        self.assertFalse(cert.root_complete)
        self.assertFalse(cert.result['native_import']['root_flags_reused'])

    def test_model_or_manifest_mismatch_rejects_before_solver_audit(self):
        for change in ['materials','body','scale','restraint','manifest']:
            raw,mesh,record=snapshot()
            if change=='materials':mesh.materials=np.array([[2.,30.,30.,18.]])
            elif change=='body':mesh.body=np.ones((1,2,2))
            elif change=='scale':mesh.stress_scale=2.
            elif change=='restraint':mesh.boundary[(0,1)]['kind']='load'
            else:record['lower_fs']=2.
            with self.subTest(change=change),patch.object(native.core,'mesh_from_arrays',return_value=mesh),patch.object(native,'evaluate') as audit,self.assertRaises(ValueError):
                native.import_side(raw,record,{},'lower')
            audit.assert_not_called()

    def test_native_success_label_cannot_accept_failed_original_physical_field(self):
        raw,mesh,record=snapshot()
        record.update(target_met=True,status='PASS')
        failure={'physical_pass':False,'witness':False,'value':1.1}
        with patch.object(native.core,'mesh_from_arrays',return_value=mesh),patch.object(native,'Prepared'),patch.object(native,'evaluate',return_value=failure):
            cert,report=native.import_side(raw,record,{},'lower')
        self.assertIsNone(cert)
        self.assertFalse(report['accepted'])


if __name__=='__main__':unittest.main()
