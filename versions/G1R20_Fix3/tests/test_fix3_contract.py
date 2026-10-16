"""Source scope and persistence wiring; actual failures are tested in owned Excel."""
from pathlib import Path
import unittest
ROOT=Path(__file__).resolve().parents[1]
BASE=ROOT.parent/'G1R20_Fix2'
class Fix3Contract(unittest.TestCase):
 def test_full_source_and_encodings(self):
  files=list((ROOT/'src/vba').glob('*.bas'))
  self.assertEqual(len(files),63)
  for p in files:
   self.assertEqual(p.read_bytes().decode('utf8'),(ROOT/'import'/p.name).read_bytes().decode('cp932'))
   self.assertNotIn(b'\n',p.read_bytes().replace(b'\r\n',b''))
  self.assertEqual((ROOT/'src/workbook/ThisWorkbook.code.txt').read_text(encoding='utf8'),(ROOT/'import/ThisWorkbook.code.txt').read_text(encoding='cp932'))
 def test_original_mathematical_kernels_unchanged(self):
  for name in ['RPX_AccuracyNative.bas','RPX_Adapt.bas','RPX_Refine.bas','RPX_FlowMesh.bas','RPX_Audit.bas','RPX_Mesh.bas','RPX_Solve.bas']:
   self.assertEqual((ROOT/'src/vba'/name).read_bytes(),(BASE/'src/vba'/name).read_bytes())
 def test_workbook_events_delegate_to_guard(self):
  text=(ROOT/'src/workbook/ThisWorkbook.code.txt').read_text(encoding='utf8')
  for a,b in [('Workbook_Open','RPX_WorkbookOpened'),('Workbook_SheetChange','RPX_WorkbookSheetChanged'),('Workbook_SheetCalculate','RPX_WorkbookSheetCalculated')]:
   self.assertIn(a,text);self.assertIn(b,text)
 def test_result_commit_precedes_current_flag(self):
  text=(ROOT/'src/vba/RPX_Control.bas').read_text(encoding='utf8')
  begin=text.index('RPX_AssertInputs "result_commit"')
  self.assertLess(text.index('Call RPX_InputCommit',begin),text.index('RPX_ResultCurrent = True',begin))
 def test_idle_observation_has_no_solve_epoch(self):
  text=(ROOT/'src/vba/RPX_Guard.bas').read_text(encoding='utf8')
  observe=text.split('Public Sub RPX_InputObserve()',1)[1].split('End Sub',1)[0]
  self.assertIn('inputActive = False',observe)
  self.assertNotIn('RPX_InputEpoch =',observe)
  self.assertNotIn('inputActive = True',observe)
 def test_saved_result_without_proof_is_marked(self):
  text=(ROOT/'src/vba/RPX_Guard.bas').read_text(encoding='utf8')
  self.assertIn('保存結果（入力との対応未確認・再解析してください）',text)
  self.assertIn('context <> RPX_CurrentInputContext()',text)
  self.assertIn('RPX_SourceRecordWrite(MAIN_SOURCE_NAMESPACE',text)
 def test_water_contract(self):
  import check_water_source_contract
  check_water_source_contract.main()
if __name__=='__main__':unittest.main()
