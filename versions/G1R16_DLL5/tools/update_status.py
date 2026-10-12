"""Compact evidence-backed task state; publication verification is separate."""
import json
from prepare import ROOT,write,sha

def main():
 def read(name):
  p=ROOT/'results'/name
  return json.loads(p.read_text(encoding='utf8')) if p.exists() else None
 core=read('core_pipeline.json');final=read('pipeline.json');delivery=read('delivery_verification.json')
 state='VALIDATED_READY_FOR_PUBLICATION' if delivery and delivery.get('reopen_whole_project_compile')=='PASS' and final and final['state']=='completed' else 'VALIDATING'
 write(ROOT/'WORK_STATUS.json',{'state':state,'protected_original_files':1085,'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'core_state':core['state'] if core else 'NOT_RUN','final_state':final['state'] if final else 'NOT_RUN','user_selected_scope':'Full levee; fixed P06A/P06N; coarse D03N/D07A; P06A first15 Fs>=0.125; coarse P06N. No full P06/tiny-Fs/26-case restart.','publication_record':'results/publication_verification.json, generated only after successful push/download verification.'})
if __name__=='__main__':main()
