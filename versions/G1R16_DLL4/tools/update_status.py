"""Record current task state without treating canceled/full-unrun cases as passes."""
import json
from prepare import ROOT,write,sha

def read(name):
 p=ROOT/'results'/name
 return json.loads(p.read_text(encoding='utf-8')) if p.exists() else None

def main():
 pipeline=read('pipeline.json');p06=read('p06a_root_pipeline.json')
 verified=read('delivery_verification.json');packaged=read('package_verification.json')
 release_ready=bool(verified and verified.get('reopen_whole_project_compile')=='PASS' and read('comparison.json') and read('root_regressions.json'))
 data={
  'state':'PACKAGED' if packaged else 'VALIDATED_READY_FOR_PUBLICATION' if release_ready else 'VALIDATING',
  'source_protected':533,
  'latest_dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),
  'scope':'DLL4 normal IPM and HSD numerical jobs; compare actual VBA/DLL3 precision, robustness and DLL3 speed; publish verified paired XLSM/DLL and evidence.',
  'user_selected_P06_scope':'Cancel full P06A, complete both-case comparison first. Then stop P06A extreme-Fs coarse exploration and compare completed practical Fs points (Fs>=0.125,15 points). P06N full adaptive run not started.',
  'P06A_full':read('p06a_full_dll4/cancellation.json'),
  'primary_comparisons':pipeline,
  'P06A_coarse_comparisons':p06,
  'tests_with_evidence':[
   'native C/C++/ASM build and native Excel import/compile',
   'HSD 6; actual VBA/DLL4 with recorded actual DLL3 three-way output',
   'LU 360:352 bitwise exact + 8 matching singular rejections',
   'NT 224:174 bitwise exact + 50 matching rejections',
   'bulk160:1239592 exact bytes, ordering68 including n>20000',
   'boundary5070: exact outputs and error partial state',
   'job controls12: cancel/OOM/subscript/unknown/input/success',
   'large262144-variable job lifecycle and real yielding',
   'mode switch2: factor invalidation and new solve',
   'C++ allocation and unknown exception injection builds/runs',
   'Issue1/2 assertions51 and native/missing/oldDLL2/oldDLL3 fallback',
   'seep convergence, nonconvergence and water-input checks'],
  'delivery_verification':verified,
  'package_verification':packaged,
  'not_completed':['P06A full adaptive search (user canceled)',
                   'P06A coarse failure reproduction below Fs0.125 in DLL4 (user narrowed extra comparison)',
                   'P06N full adaptive search (not started)',
                   'full25/26 case batch (not resumed)'],
  'numerical_scope':'Nonlinear loop control, Fs search, physical audits, rescue policy and mesh remain VBA. Linear solve/refinement jobs and fused numerical loops execute in DLL.',
  'snapshot_semantics':'Immutable validation snapshot taken before ZIP creation and GitHub publication. Post-package and post-upload checks are recorded separately, so this file remains covered by the archive manifest.',
  'publication':'GitHub upload and HTTP/SHA256 download verification follow this snapshot. verify_published.py records the actual result separately; the delivery message reports the verified published commit.'
 }
 write(ROOT/'WORK_STATUS.json',data)

if __name__=='__main__':main()
