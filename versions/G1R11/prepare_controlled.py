from pathlib import Path
R=Path(__file__).resolve().parent;OLD=Path('C:/Users/link_/Desktop/RPFEM_G1R10_Mesh_20261007')
s=(OLD/'native_controlled.py').read_text(encoding='utf-8')
s=s.replace('from native_mesh import *','from native_common import *')
s=s.replace("shutil.copy2(source['path'],work)",'shutil.copy2(source,work)')
s=s.replace("setvalue(book,'ADAPT_MESH_POLICY','ROBUST_REFINE')","setvalue(book,'ADAPT_MESH_POLICY','LEGACY_THEN_REFINE');setvalue(book,'ADAPT_REFINE_CAP',1000);setvalue(book,'FS_BRACKET_POLICY','AUDITED_UPPER_LIMIT')")
s=s.replace("R/'results/kinematic_scores.npz'","R/'results/staged_1000_selection.npz'")
s=s.replace("['neighbour']","['score']")
# Probe the new indicator using the same q4 upper field as the Python-first candidate.
a=s.index("    ed=cells('xl/worksheets/sheet8.xml')");b=s.index("    (R/'results/velocity.bin')",a)
s=s[:a]+"    v=np.load(R/'results/staged_1000_selection.npz')['velocity'].reshape(781,12)\n"+s[b:]
# Isolate scheduling and failure recovery. Band geometry itself remains byte-unchanged.
s=s.replace("    text=replaceproc(text,'Private Sub RestoreFinalBounds(','End Sub',final)","    text=replaceproc(text,'Private Sub RestoreFinalBounds(','End Sub',final)\n    text=replaceproc(text,'Private Sub RPX_BuildBandMesh(','End Sub','Private Sub RPX_BuildBandMesh(ByRef baseline As RPX_MeshState)\\n    RPX_RestoreState baseline\\nEnd Sub')")
s=s.replace("'adapt_1_refine'","'staged_refine'")
s=s.replace("<=1536","<=1000")
s=s.replace("assert 'status=aborted' in out and 'elements=601;' in out,out","assert 'status=staged_candidate_rejected' in out,out")
s=s.replace("(R/'results/native_controlled.json')","(R/'results/native_controlled.json')")
s=s.replace("RUNTIME_SCOPE='NATIVE_CONTROL_FLOW_WITH_TEST_ONLY_MOCK_EVALUATE_NOT_PHYSICAL_SOLVE'","RUNTIME_SCOPE='NATIVE_CONTROL_FLOW_WITH_TEST_ONLY_MOCK_EVALUATE_AND_BAND_NOT_PHYSICAL_SOLVE'")
(R/'native_controlled.py').write_text(s,encoding='utf-8')
print('Prepared controlled scheduling/failure tests; mock evaluation and band are test-only')
