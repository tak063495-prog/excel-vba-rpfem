"""Mirror native target semantics without changing solver tolerances.

RPX_Fs stops on width <= FS_TOL * max(1, bracket midpoint).
RPX_Guard.RPX_TargetMet explicitly rejects comparing the absolute maximum root
width directly to FS_TOL. Saved root-valid/completion flags come from the native
per-owner bracket checks; independent field audits are still mandatory.
"""
def grade(record):
    r=dict(record)
    r['native_worker_status']=record.get('status')
    if not r.get('fs_obtained'):return r
    flags=r.get('summary_fields',[])
    roots_complete=len(flags)>=16 and flags[12:14]==['True','True'] and flags[14:16]==['False','False']
    gap_target=float(r.get('actual_settings',{}).get('GAP',.01))
    target=bool(r.get('independent_audit',{}).get('pass') and r.get('result_current') and roots_complete and not r.get('root_search_incomplete') and r['relative_gap']<=gap_target)
    r['target_met']=target
    r['status']='PASS_TARGET' if target else 'FS_AVAILABLE_PARTIAL'
    r['global_incomplete_note']=flags[6] if len(flags)>6 else ''
    reasons=[]
    if r['relative_gap']>gap_target:reasons.append('GAP_TARGET_NOT_MET')
    if len(flags)<16:reasons.append('ROOT_OWNER_FLAGS_MISSING')
    else:
        if flags[12]!='True':reasons.append('LOWER_ROOT_NOT_VALID')
        if flags[13]!='True':reasons.append('UPPER_ROOT_NOT_VALID')
        if flags[14]!='False':reasons.append('LOWER_OWNER_SEARCH_INCOMPLETE')
        if flags[15]!='False':reasons.append('UPPER_OWNER_SEARCH_INCOMPLETE')
    if r.get('root_search_incomplete') and not reasons:reasons.append('NATIVE_GLOBAL_INCOMPLETE')
    r['target_unmet_reasons']=reasons
    r['root_tolerance_contract']='width <= FS_TOL * max(1, bracket midpoint); per-owner native root validity and completion flags'
    r['grading_basis']='RPX_Fs.bas:572-573,613; RPX_Guard.bas:372-381; independent raw field audit'
    r['grading_revision']=1
    return r
