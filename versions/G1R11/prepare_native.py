from pathlib import Path
import shutil
R=Path(__file__).resolve().parent;OLD=Path('C:/Users/link_/Desktop/RPFEM_G1R10_Mesh_20261007')
s=(OLD/'tests/MeshProbe.bas').read_text(encoding='utf-8')
s=s.replace('ByVal budget As Boolean) As String','ByVal budget As Boolean, Optional ByVal forceLegacy As Boolean = False) As String',1)
s=s.replace('Then RPX_RefineMesh w, fraction','Then RPX_RefineMesh w, fraction, forceLegacy',1)
(R/'tests/MeshProbe.bas').write_text(s,encoding='utf-8')
print('Prepared test-only native mesh probe')
