from native_repair import ROOT,OLD
import shutil
for name in ['audit_batch.py','grade_results.py']:
    s=(OLD/name).read_text(encoding='utf-8')
    if name=='audit_batch.py':
        s=s.replace('from zero_phi import audit_zero','from zero_phi import audit_zero\nfrom audit_precision import audit_precision')
        old="        raw={k:max(0.,float(v))*m.stress_scale"
        new="""        precision=None
        if side=='upper':
            m.provenance['case']=case
            precision=audit_precision(m,x,meta['fs'],'upper')
            # Record ordinary Double diagnostics; final decision uses independent 50-digit integration.
            if np.all(m.materials[:,0]==0):
                au['ordinary_double_pass']=au['pass']
                au['pass']=bool(precision['accepted'])
                au['acceptance_basis']='50-digit independent physical audit; unchanged 1e-7'
        raw={k:max(0.,float(v))*m.stress_scale"""
        assert old in s;s=s.replace(old,new)
        s=s.replace("'physical_audit':au,'raw_lower_errors'","'physical_audit':au,'precision_audit':precision,'raw_lower_errors'")
    (ROOT/'tests'/name).write_text(s,encoding='utf-8')
s=(OLD/'run_native25.py').read_text(encoding='utf-8')
s=s.replace('from prepare_inputs import ROOT,REPO,SOURCE,SOURCE_SHA',"""from prepare_inputs import REPO,SOURCE,SOURCE_SHA
ROOT=Path(__file__).resolve().parents[1]""")
s=s.replace("TEMPLATE=ROOT/'prepared_template.xlsm'","TEMPLATE=ROOT/'cases/D07A/RPFEM_20261010_G1R13_D07A.xlsm'")
s=s.replace("path=folder/f'G1R12_{case}_Full1536_Q8.xlsm'","path=folder/f'RPFEM_20261010_G1R13_{case}_Full1536_Q8.xlsm'")
s=s.replace('G1R12実Excel25ケース検証。','G1R13 D03N/D07A修復の実Excel検証。')
s=s.replace("from prepare_inputs import REPO,SOURCE,SOURCE_SHA","sys.path.insert(0,'"+str(OLD).replace('\\','/')+"')\nfrom prepare_inputs import REPO,SOURCE,SOURCE_SHA")
(ROOT/'tests/run_full.py').write_text(s,encoding='utf-8')
shutil.copytree(OLD/'inputs',ROOT/'inputs',dirs_exist_ok=True)
print('RUNNER READY')
