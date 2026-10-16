"""Scope/guard contracts complement the owned Excel state regressions."""
from pathlib import Path
import re,unittest
ROOT=Path(__file__).resolve().parents[1]
OLD=ROOT.parent/'G1R20_Fix1/src/vba';NEW=ROOT/'src/vba'
def source(root,name):return (root/name).read_text(encoding='utf8')
def proc(text,name):
    m=re.search(r'(?im)^(?:Public |Private )?(?:Sub|Function) '+re.escape(name)+r'\b.*?\nEnd (?:Sub|Function)\s*$',text,re.S)
    if not m:raise AssertionError('Missing '+name)
    return m.group(0)
class UIContract(unittest.TestCase):
    def test_setup_busy_precedes_scope_or_mutation(self):
        s=proc(source(NEW,'RPX_UI.bas'),'RPX_Setup')
        self.assertLess(s.index('If RPX_Busy Then'),s.index('oldWrite ='))
        self.assertLess(s.index('Call RPX_Invalidate'),s.index('For Each name'))
    def test_setup_default_input_and_setting_values_unchanged(self):
        old=proc(source(OLD,'RPX_UI.bas'),'RPX_Setup');new=proc(source(NEW,'RPX_UI.bas'),'RPX_Setup')
        def body(s):return s[s.index('    For Each name'):s.index('    RPX_InternalWrite =',s.index('Call RPX_WriteSlopeInputs'))]
        self.assertEqual(body(old),body(new))
    def test_setup_error_preserved_and_nested_scope_restored(self):
        s=proc(source(NEW,'RPX_UI.bas'),'RPX_Setup').split('Failed:',1)[1]
        self.assertLess(s.index('number = err.number'),s.index('Call RPX_Invalidate'))
        self.assertLess(s.index('Resume FailedCleanup'),s.index('Call RPX_Invalidate'))
        self.assertIn('RPX_InternalWrite = oldWrite',s)
        self.assertIn('err.Raise number, source, message',s)
    def test_mesh_failure_tracks_work_and_preserves_primary_error(self):
        s=proc(source(NEW,'RPX_Control.bas'),'RPX_CreateMesh')
        self.assertLess(s.index('Call RPX_Invalidate'),s.index('workEntered = True'))
        self.assertLess(s.index('workEntered = True'),s.index('Call StartWork'))
        fail=s.split('Failed:',1)[1]
        self.assertLess(fail.index('RPX_LastFailureNumber = err.number'),fail.index('On Error Resume Next'))
        self.assertLess(fail.index('Resume FailedCleanup'),fail.index('On Error Resume Next'))
        self.assertIn('RPX_MeshCurrent = False: RPX_ResultCurrent = False',fail)
        self.assertIn('If workEntered Then Call EndWork',fail)
    def test_cleanup_remains_under_error_suppression_until_cancel_restored(self):
        s=proc(source(NEW,'RPX_Control.bas'),'EndWork')
        self.assertLess(s.index('On Error Resume Next'),s.index('RPX_SchurRelease'))
        self.assertLess(s.index('Application.EnableCancelKey = priorCancelKey'),s.index('On Error GoTo 0'))
    def test_other_ui_and_control_procedures_unchanged(self):
        for name,changed in [('RPX_UI.bas',['RPX_Setup','RPX_AccuracyPreset','RPX_Speed2pctPreset','RPX_G20ImprovementPreset']),('RPX_Control.bas',['EndWork','RPX_CreateMesh','RPX_Run'])]:
            old=source(OLD,name);new=source(NEW,name)
            for p in changed:old=old.replace(proc(old,p),'');new=new.replace(proc(new,p),'')
            self.assertEqual(old,new,name)
    def test_presets_only_deactivate_active_error_handlers(self):
        for name in ['RPX_AccuracyPreset','RPX_Speed2pctPreset','RPX_G20ImprovementPreset']:
            old=proc(source(OLD,'RPX_UI.bas'),name);new=proc(source(NEW,'RPX_UI.bas'),name)
            self.assertEqual(old,new.replace('    Resume FailedCleanup\nFailedCleanup:\n',''))
    def test_run_numerical_and_failure_body_unchanged_except_resume(self):
        old=proc(source(OLD,'RPX_Control.bas'),'RPX_Run');new=proc(source(NEW,'RPX_Control.bas'),'RPX_Run')
        self.assertEqual(old,new.replace('    Resume FailedCleanup\nFailedCleanup:\n',''))
if __name__=='__main__':unittest.main()
