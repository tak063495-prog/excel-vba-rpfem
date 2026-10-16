"""Small independent model-consistency rejection tests; no numerical solves."""
from pathlib import Path
import importlib.util
import tempfile
import unittest
import zipfile

spec=importlib.util.spec_from_file_location('water_auditor',Path(__file__).with_name('audit_cold_water_native.py'))
audit=importlib.util.module_from_spec(spec);spec.loader.exec_module(audit)
np=audit.np


class WaterAuditTests(unittest.TestCase):
    def fixture(self):
        xy=np.array([[0.,2.],[1.,2.],[0.,3.],[0.,3.],[1.,4.],[0.,4.]])
        tri=np.array([[0,1,2],[3,4,5]])
        mats=np.array([[10.,25.,25.,20.]])
        ids=np.array([0,0]);body=np.zeros((2,2,2))
        body[:,1,1]=[-20.+9.81,-20.]
        return xy,tri,mats,ids,body

    def test_common_hydrostatic_piecewise_force(self):
        f=self.fixture();np.testing.assert_array_equal(audit.common_body(*f),f[-1])

    def test_reject_dry_or_double_added_or_wrong_phase_body(self):
        for change in ['dry','double','fixed']:
            with self.subTest(change=change):
                xy,tri,mats,ids,body=self.fixture()
                if change=='dry':body[0,1,1]=-20.
                elif change=='double':body[0,1,1]+=9.81
                else:body[0,1,0]=body[0,1,1];body[0,1,1]=0.
                with self.assertRaises(AssertionError):audit.common_body(xy,tri,mats,ids,body)

    def test_reject_waterline_crossing_despite_matching_average(self):
        xy,tri,mats,ids,body=self.fixture();xy[0,1]=4.
        with self.assertRaisesRegex(ValueError,'crosses'):audit.common_body(xy,tri,mats,ids,body)

    def test_internal_waterline_constraint_is_geometry_only(self):
        xy=np.array([[0.,3.],[1.,3.],[1.,4.]])
        self.assertEqual(audit.constraint_kind((0,1),xy,{}),'waterline_geometric')
        self.assertEqual(audit.constraint_kind((1,2),xy,{(1,2):{}}),'exterior')
        with self.assertRaisesRegex(ValueError,'not on'):audit.constraint_kind((0,2),xy,{})

    def workbook(self,path,bad_head=False,bad_pressure=False):
        rows=[]
        points=[(0.,2.),(1.,3.),(0.,4.)]
        for i,(x,y) in enumerate(points,2):
            values={'A':i-1,'B':x,'C':y,'D':4. if bad_head and i==2 else 3.,
                    'E':3.-y,'F':9.81*(3.-y)+(1. if bad_pressure and i==2 else 0.)}
            rows.append('<row r="%d">'%i+''.join('<c r="%s%d"><v>%s</v></c>'%(k,i,v) for k,v in values.items())+'</row>')
        rows+=['<row r="2"><c r="K2" t="b"><v>1</v></c></row>',
               '<row r="7"><c r="K7"><v>3</v></c></row>']
        with zipfile.ZipFile(path,'w') as archive:
            archive.writestr('xl/workbook.xml','<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="浸透結果" sheetId="1" r:id="rId1"/></sheets></workbook>')
            archive.writestr('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
            archive.writestr('xl/worksheets/sheet1.xml','<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+''.join(rows)+'</sheetData></worksheet>')

    def test_saved_head_and_signed_pressure_all_nodes(self):
        summary={'prescribed_water':{'result_K2_K10':[[True],[],[],[],[],[3]]}}
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'synthetic.xlsm'
            self.workbook(path)
            self.assertTrue(audit.workbook_water(path,summary)['all_saved_node_heads_exactly_3'])
            self.workbook(path,bad_head=True)
            with self.assertRaisesRegex(ValueError,'head differs'):audit.workbook_water(path,summary)
            self.workbook(path,bad_pressure=True)
            with self.assertRaisesRegex(ValueError,'pressure differs'):audit.workbook_water(path,summary)


if __name__=='__main__':unittest.main()
