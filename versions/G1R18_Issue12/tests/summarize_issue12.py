"""Summarize Excel run folders (result.json + RPFEM_logs/*analysis.tsv) into JSON/Markdown rows."""
import json,glob,os,sys
root=sys.argv[1]; rows=[]
for d in sorted(os.listdir(root)):
    p=os.path.join(root,d)
    if not os.path.isfile(os.path.join(p,'result.json')): continue
    j=json.loads(open(os.path.join(p,'result.json'),encoding='utf-8-sig').read())
    st=dict(x.split('=',1) for x in (j.get('state') or '').split(';') if '=' in x)
    logs=glob.glob(os.path.join(p,'RPFEM_logs','*analysis.tsv'))
    F=[l.split('\t') for l in open(logs[0],encoding='utf-16').read().splitlines()] if logs else []
    F=[f for f in F if len(f)>10]
    fin=[f for f in F if f[3]=='final']
    r=dict(run=d,status=j.get('status'),error=j.get('error'),lower=j.get('lower'),upper=j.get('upper'),seconds=j.get('seconds'),
      cache_hits=sum(1 for f in F if f[6]=='lower_cache_lookup' and 'hit=1' in f[10]),
      final_cache_hits=sum(1 for f in fin if f[6]=='lower_cache_lookup' and 'hit=1' in f[10]),
      evidence_invalidations=[f[10] for f in F if f[6]=='lower_cache_lookup' and 'original_evidence' in f[10]],
      restore_details=[f[10] for f in F if f[6]=='lower_cache_restore'],
      final_hsd_not_converged=sum(1 for f in fin if f[6]=='error_origin' and 'HSD_NOT_CONVERGED' in f[10]),
      final_rejected=[f[10] for f in fin if f[6]=='final_candidate_rejected'],
      prepared_drops=sum(1 for f in F if f[6]=='lower_prepared_drop'),
      run_end=[f[10].split(';')[0] for f in F if f[6]=='run_end'],
      lower_q=st.get('lower_q'),upper_q=st.get('upper_q'),lower_root=st.get('lower_root'),upper_root=st.get('upper_root'),
      search_incomplete=st.get('search_incomplete'),final_original_q=st.get('original_q'),source_sha256=j.get('source_sha256'),output_sha256=j.get('output_sha256'),settings=j.get('settings'))
    rows.append(r)
json.dump(rows,open(sys.argv[2],'w',encoding='utf-8'),ensure_ascii=False,indent=1)
for r in rows:
    print(f"{r['run']:24s} {r['status']} L={r['lower']!r} U={r['upper']!r} hits={r['cache_hits']} fhit={r['final_cache_hits']} inval={len(r['evidence_invalidations'])} fHSD={r['final_hsd_not_converged']} rej={len(r['final_rejected'])} q={r['lower_q']}/{r['upper_q']} root={r['lower_root']}/{r['upper_root']} inc={r['search_incomplete']} oq={r['final_original_q']} {r['run_end']} t={r['seconds'] and round(r['seconds'])}")
