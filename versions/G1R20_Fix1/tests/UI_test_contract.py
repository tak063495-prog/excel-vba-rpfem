"""Solver-free contracts for the public preset fix; actual Excel is separate."""
from pathlib import Path
import re,unittest
ROOT=Path(__file__).resolve().parents[1]
BASE=ROOT.parent/'G1R20'
def procedures(text):
    return {m.group(1):m.group(0) for m in re.finditer(r'Public Sub (RPX_\w+Preset)\(\).*?^End Sub',text,re.M|re.S)}
class PresetContract(unittest.TestCase):
    def test_cp932_exact(self):
        self.assertEqual((ROOT/'src/vba/RPX_UI.bas').read_text(encoding='utf8'),(ROOT/'import/RPX_UI.bas').read_text(encoding='cp932'))
    def test_only_three_public_presets_changed(self):
        old=(BASE/'src/vba/RPX_UI.bas').read_text(encoding='utf8');new=(ROOT/'src/vba/RPX_UI.bas').read_text(encoding='utf8')
        a,b=procedures(old),procedures(new)
        for name in ['RPX_AccuracyPreset','RPX_Speed2pctPreset','RPX_G20ImprovementPreset']:
            self.assertNotEqual(a[name],b[name]);old=old.replace(a[name],'');new=new.replace(b[name],'')
        self.assertEqual(old,new,'Changes outside the authorized three preset routines')
    def test_every_public_preset_has_guarded_scope_and_original_error(self):
        text=(ROOT/'src/vba/RPX_UI.bas').read_text(encoding='utf8')
        for name,body in procedures(text).items():
            self.assertLess(body.index('If RPX_Busy Then'),body.index('oldWrite = RPX_InternalWrite'),name)
            self.assertLess(body.index('On Error GoTo Failed'),body.index('RPX_InternalWrite = True'),name)
            self.assertEqual(body.count('RPX_InternalWrite = oldWrite'),2,name)
            self.assertEqual(body.count('If Not oldWrite Then Call RPX_Invalidate'),2,name)
            self.assertIn('number = err.number: message = err.description: source = err.source',body)
            self.assertIn('err.Raise number, source, message',body)
    def test_migration_initialization_precedes_first_setting_write(self):
        body=procedures((ROOT/'src/vba/RPX_UI.bas').read_text(encoding='utf8'))['RPX_AccuracyPreset']
        self.assertLess(body.index('Call RPX_EnsureUpperPilotSetting'),body.index('AccuracySet'))
        self.assertLess(body.index('Call RPX_EnsureG1R18Settings'),body.index('AccuracySet'))
    def test_prescribed_setting_values_unchanged(self):
        old=procedures((BASE/'src/vba/RPX_UI.bas').read_text(encoding='utf8'));new=procedures((ROOT/'src/vba/RPX_UI.bas').read_text(encoding='utf8'))
        for name in old:
            self.assertEqual(re.findall(r'AccuracySet[^\n]*',old[name]),re.findall(r'AccuracySet[^\n]*',new[name]),name)
if __name__=='__main__':unittest.main()
