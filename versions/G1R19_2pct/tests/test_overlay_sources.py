"""Byte parity and unchanged original acceptance gates for the four-module overlay."""
from pathlib import Path
import importlib.util, json, unittest

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('overlay_sources',ROOT/'tools/current_sources.py')
sources=importlib.util.module_from_spec(spec);spec.loader.exec_module(sources)

def body(text,name):
    pos=text.index(name)
    start=text.rfind('\n',0,pos)+1
    kind='Function' if 'Function ' in text[start:text.index('\n',pos)] else 'Sub'
    return text[start:text.index('End '+kind,pos)+len('End '+kind)]

class OverlayIntegrity(unittest.TestCase):
    def test_complete_import_pairs_and_manifest(self):
        paths=sources.module_paths()
        self.assertEqual(len(paths),62)
        metadata=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
        for name,row in metadata['modules'].items():
            self.assertEqual(row['utf8_sha256'],sources.sha(ROOT/'src/vba'/name))
            self.assertEqual(row['cp932_sha256'],sources.sha(ROOT/'import'/name))

    def test_original_scores_and_terminal_validation_unchanged(self):
        current=(ROOT/'src/vba/RPX_Adapt.bas').read_text(encoding='utf8')
        release=(sources.BASE/'src/vba/RPX_Adapt.bas').read_text(encoding='utf8')
        issue=(ROOT.parent/'G1R19_Issue11/src/vba/RPX_Adapt.bas').read_text(encoding='utf8')
        for name in ['ComputeShearWeight(', 'ComputeKinematicRefineWeight(', 'RecheckOwnQuadrature(',
                     'TerminalOwnMetadataReady(', 'TerminalOwnFieldReady(']:
            # The definitions follow any call sites; match the declaration prefix.
            token=('Private Function ' if name.startswith(('Recheck','Terminal')) else 'Private Sub ')+name
            self.assertEqual(body(current,token),body(release,token))
        token='Public Function RPX_TryTerminalPairReuse('
        self.assertEqual(body(current,token),body(issue,token))
        for name in ['RPX_Guard','RPX_Certificate','RPX_AccuracyMesh']:
            self.assertEqual(sources.module_paths()[name],sources.BASE/'src/vba'/(name+'.bas'))
        for module,tokens in {
            'RPX_Fs':['Public Function RPX_ActualRootBracket(', 'Private Function TryFsTrialPoint(', 'Private Sub SaveBracket('],
            'RPX_AccuracyNative':['Public Function RPX_AccuracyExport(', 'Private Function CurrentProgram(',
                                  'Public Function RPX_AccuracyPrimalAudit(', 'Public Function RPX_AccuracyDualAudit(']
        }.items():
            updated=sources.module_paths()[module].read_text(encoding='utf8')
            baseline=(sources.BASE/'src/vba'/(module+'.bas')).read_text(encoding='utf8')
            for token in tokens:self.assertEqual(body(updated,token),body(baseline,token))

if __name__=='__main__':unittest.main()
