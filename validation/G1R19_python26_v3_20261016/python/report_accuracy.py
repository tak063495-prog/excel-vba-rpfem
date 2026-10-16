"""Freeze and re-audit the strongest same-model bounds across experiment roots.

Run after producers finish. Mutable or inconsistent source snapshots are
skipped and recorded, never silently treated as completed accuracy goals.
"""
from __future__ import annotations
import argparse
import csv
import hashlib
import io
import json
from pathlib import Path
from zipfile import BadZipFile

import accuracy_driver as driver
from experiment_core import np,Prepared,evaluate,write

DEFAULT_TABLE=driver.REPO/'docs/benchmarks/G1R17_native26_20261013/data/RESULTS.csv'


def digest(data):return hashlib.sha256(data).hexdigest()


def root_is_complete(metadata,mesh,fs,mode,case,tolerance=1e-4):
    """The modern driver's identity/sign/endpoint requirements, without I/O."""
    root=metadata.get('audit',{}).get('discrete_root_search',{})
    bracket=root.get('bracket',[])
    ordered=bool(isinstance(bracket,list) and len(bracket)==2 and driver.finite(*bracket) and 0<bracket[0]<bracket[1])
    narrow=bool(ordered and bracket[1]-bracket[0]<=tolerance*max(1.,sum(bracket)/2))
    identity=bool(metadata.get('fs')==fs and metadata.get('mesh_id')==mesh.mesh_id and
        metadata.get('physical_model_id')==mesh.physical_model_id and metadata.get('case')==case and
        metadata.get('mode')==mode and metadata.get('q')==8 and root.get('mesh_id')==mesh.mesh_id and
        root.get('physical_model_id')==mesh.physical_model_id and root.get('case')==case and
        root.get('policy')==mesh.provenance['psi_policy'] and root.get('q')==8 and root.get('mode')==mode)
    return bool(identity and narrow and root.get('bracket_signs')==[1,-1] and
        metadata.get('root_complete') is True and root.get('root_complete') is True and
        root.get('endpoint')==fs and fs==(bracket[0] if mode=='lower' else bracket[1]))


def inside_file(folder,name):
    path=(folder/name).resolve()
    if not path.is_relative_to(folder.resolve()):
        raise ValueError('A certificate file points outside its case folder')
    return path


def read_tracked(path,tracked):
    data=path.read_bytes();tracked[path]=digest(data);return data


def unchanged(tracked):
    try:return all(digest(p.read_bytes())==h for p,h in tracked.items())
    except OSError:return False


def collect_candidate(meta_path,mode,case,rec,roots,tolerance):
    folder=meta_path.parent;tracked={}
    catalog_path=driver.CATALOG/f'{case}.json'
    if json.loads(read_tracked(catalog_path,tracked).decode('utf8'))!=rec:
        raise ValueError('MUTABLE_INPUT_CHANGED_BEFORE_REAUDIT')
    document=json.loads(read_tracked(meta_path,tracked).decode('utf8'))
    metadata=document if document.get('mode')==mode and 'fs' in document else document.get('certificates',document).get(mode)
    if metadata is None:return None
    mesh_path=inside_file(folder,metadata.get('mesh_path',f'best_{mode}_mesh.npz'))
    field_path=inside_file(folder,metadata.get('field_path',f'best_{mode}_field.npz'))
    with np.load(io.BytesIO(read_tracked(mesh_path,tracked)),allow_pickle=False) as data:
        mesh=driver.core.mesh_from_arrays(data['xy'].copy(),data['tri'].copy(),data['material_id'].copy(),rec)
    with np.load(io.BytesIO(read_tracked(field_path,tracked)),allow_pickle=False) as data:
        if str(data['mode'])!=mode or int(data['q'])!=8:
            raise ValueError('Stored mode or quadrature differs')
        field=data['x'].copy();fs=float(data['fs'])
    if not driver.finite(fs) or fs<=0:
        raise ValueError('Stored Fs is nonfinite or nonpositive')
    if metadata.get('fs')!=fs or metadata.get('mesh_id')!=mesh.mesh_id or metadata.get('physical_model_id')!=mesh.physical_model_id:
        raise ValueError('Metadata, mesh, model, and field Fs are inconsistent')
    # The combined seed may refer to the original import's root record. Read
    # it only inside a supplied source tree or the repository, tracking bytes.
    root=metadata.get('audit',{}).get('discrete_root_search',{})
    if not root:
        source=Path(metadata.get('source','')).resolve()
        allowed=source.is_relative_to(driver.REPO.resolve()) or any(source.is_relative_to(r) for r in roots)
        imported=source/'import_audit.json'
        if allowed and imported.is_file():
            old=json.loads(read_tracked(imported,tracked).decode('utf8')).get('certificates',{}).get(mode,{})
            if old.get('fs')==fs and old.get('mesh_id')==mesh.mesh_id:
                metadata=old
                root=old.get('audit',{}).get('discrete_root_search',{})
    result=evaluate(Prepared(mesh,fs,mode,8,reduction='none'),field)
    if not unchanged(tracked):raise ValueError('MUTABLE_SOURCE_CHANGED_DURING_REAUDIT')
    if not result['physical_pass'] or not result['witness']:
        raise ValueError('The unchanged original physical witness gate rejected the field')
    if root:result['discrete_root_search']=root
    complete=root_is_complete(metadata,mesh,fs,mode,case,tolerance)
    return {'case':case,'mode':mode,'fs':fs,'mesh':mesh,'field':field,'audit':result,
            'root_complete':complete,'source':str(meta_path),'tracked':tracked,
            'source_hashes':{str(p):h for p,h in tracked.items()},
            'root_incomplete_reason':None if complete else 'Modern root identity, signs, endpoint, or width requirements are not satisfied'}


def source_files(roots,exclude=None):
    files=set()
    for root in roots:
        if root.is_file() and root.name in ['summary.json','import_audit.json','best_lower_certificate.json','best_upper_certificate.json']:files.add(root)
        elif root.is_dir():
            files.update(root.rglob('summary.json'));files.update(root.rglob('import_audit.json'))
            files.update(root.rglob('best_lower_certificate.json'));files.update(root.rglob('best_upper_certificate.json'))
    return sorted(p for p in files if exclude is None or not p.is_relative_to(exclude))


def case_of(path):
    try:
        d=json.loads(path.read_text(encoding='utf8'))
        return d.get('case') or next((c.get('case') for c in d.get('certificates',{}).values() if isinstance(c,dict)),None) or path.parent.name
    except (OSError,ValueError):return path.parent.name


def freeze_candidate(candidate,folder):
    # Recheck source identities after all candidates have been evaluated.
    if not unchanged(candidate['tracked']):raise ValueError('MUTABLE_SOURCE_CHANGED_BEFORE_FREEZE')
    side=candidate['mode'];folder.mkdir(parents=True,exist_ok=True)
    cert=driver.Certificate(side,candidate['fs'],candidate['mesh'],candidate['field'],
                            candidate['audit'],candidate['source'],candidate['root_complete'])
    metadata=cert.publish(folder)
    metadata.update(policy=candidate['mesh'].provenance['psi_policy'],source_hashes=candidate['source_hashes'])
    write(folder/f'best_{side}_certificate.json',metadata)
    return metadata


def number(value):
    try:
        v=float(value)
        return v if np.isfinite(v) else None
    except (TypeError,ValueError):return None


def run(sources,output,table=DEFAULT_TABLE,tolerance=1e-4):
    roots=[p.resolve() for p in sources];output=output.resolve()
    if any(output==r or r.is_relative_to(output) for r in roots):
        raise ValueError('The report output cannot contain or replace its source roots')
    baseline=list(csv.DictReader(table.open(encoding='utf-8-sig',newline='')))
    if not baseline:raise ValueError('The baseline table has no cases')
    by_case={row['case']:row for row in baseline};files=source_files(roots,output)
    catalog={case:json.loads((driver.CATALOG/f'{case}.json').read_text(encoding='utf8')) for case in by_case}
    candidates={case:{'lower':[],'upper':[]} for case in by_case};issues=[];valid_records=[]
    for path in files:
        case=case_of(path)
        if case not in by_case:continue
        try:
            rec=catalog[case]
            if rec['source_case'].get('pore'):
                raise ValueError('COUPLED_PORE_SNAPSHOT_REQUIRED: dry reconstruction is prohibited')
            for mode in ['lower','upper']:
                try:
                    candidate=collect_candidate(path,mode,case,rec,roots,tolerance)
                    if candidate is not None:
                        candidates[case][mode].append(candidate)
                        valid_records.append({k:candidate[k] for k in ['case','mode','fs','source','root_complete','root_incomplete_reason']})
                except (OSError,ValueError,KeyError,AssertionError,TypeError,IndexError,RuntimeError,BadZipFile) as err:
                    issues.append({'case':case,'mode':mode,'source':str(path),'reason':str(err)})
        except (OSError,ValueError,KeyError) as err:
            issues.append({'case':case,'source':str(path),'reason':str(err)})
    output.mkdir(parents=True,exist_ok=True);snapshots=output/'snapshots';comparison=[];case_reports=[]
    for case,old in by_case.items():
        chosen={}
        for mode in ['lower','upper']:
            ordered=sorted(candidates[case][mode],key=lambda c:(-c['fs'] if mode=='lower' else c['fs'],not c['root_complete'],c['source']))
            for candidate in ordered:
                try:
                    chosen[mode]=freeze_candidate(candidate,snapshots/case);break
                except (OSError,ValueError) as err:
                    issues.append({'case':case,'mode':mode,'source':candidate['source'],'reason':str(err)})
        status='COUPLED_PORE_SNAPSHOT_REQUIRED' if catalog[case]['source_case'].get('pore') else 'NO_REAUDITED_PAIR'
        report={'case':case,'target_met':False,'root_complete':False,'status':status,
                'lower':chosen.get('lower'),'upper':chosen.get('upper'),'relative_width_symmetric':None}
        if set(chosen)=={'lower','upper'}:
            lo,hi=chosen['lower'],chosen['upper']
            if lo['physical_model_id']!=hi['physical_model_id'] or not 0<lo['fs']<=hi['fs']:
                report['status']='INCONSISTENT_SAME_MODEL_BOUNDS'
                issues.append({'case':case,'reason':'Chosen extrema reverse or belong to different physical models'})
            else:
                width=2*(hi['fs']-lo['fs'])/(lo['fs']+hi['fs']);complete=lo['root_complete'] and hi['root_complete']
                met=bool(width<=.01 and complete)
                report.update(relative_width_symmetric=width,root_complete=complete,target_met=met,
                              status='VERIFIED_TARGET' if met else 'ROOT_INCOMPLETE' if not complete else 'GAP_TARGET_NOT_MET')
        write(snapshots/case/'summary.json',report);case_reports.append(report)
        comparison.append({'case':case,'old_status':old.get('status'),'old_lower_fs':number(old.get('lower_fs')),
            'old_upper_fs':number(old.get('upper_fs')),'old_gap_percent':number(old.get('gap_percent')),
            'old_target_met':old.get('target_met')=='True','new_status':report['status'],
            'new_lower_fs':chosen.get('lower',{}).get('fs'),'new_upper_fs':chosen.get('upper',{}).get('fs'),
            'new_gap_percent':None if report['relative_width_symmetric'] is None else 100*report['relative_width_symmetric'],
            'new_root_complete':report['root_complete'],'new_target_met':report['target_met'],
            'lower_elements':chosen.get('lower',{}).get('elements'),'upper_elements':chosen.get('upper',{}).get('elements'),
            'lower_source':chosen.get('lower',{}).get('source'),'upper_source':chosen.get('upper',{}).get('source')})
    with (output/'COMPARISON_26.csv').open('w',encoding='utf-8-sig',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=list(comparison[0]));writer.writeheader();writer.writerows(comparison)
    previously_unmet=[c for c in case_reports if by_case[c['case']].get('target_met')!='True']
    payload={'scope':'Independently re-audited floating-tolerance witnesses; latest strict root metadata required',
             'fs_tolerance':tolerance,'sources':[str(p) for p in roots],'cases':case_reports,
             'implementation_hashes':driver.SOURCE_HASHES,
             'report_code_hash':digest(Path(__file__).read_bytes()),'baseline_table_hash':digest(table.read_bytes()),
             'candidate_audits':valid_records,'skipped_or_rejected':issues,
             'all_26_targets_revalidated':len(case_reports)==26 and all(c['target_met'] for c in case_reports),
             'previously_unmet_count':len(previously_unmet),
             'previously_unmet_verified_count':sum(c['target_met'] for c in previously_unmet),
             'all_previously_unmet_targets_verified':bool(previously_unmet and all(c['target_met'] for c in previously_unmet))}
    write(output/'rollup.json',payload)
    def fmt(v,d=6):return '—' if v is None else f'{v:.{d}f}'
    lines=['# RPFEM 精度改善の再監査報告（2026-10-13）','',
           '最良下界と最良上界を、それぞれ元のメッシュと場で再監査して保存した。',
           '上下界幅は 2 × (上界 − 下界) / (上界 + 下界)。幅1%以下に加え、両側の最終Fs根区間10⁻⁴と、case・mode・q・policy・model・mesh・端点・符号の同一性を要求する。',
           '数値最適化の反対端点は連続体の追加境界として扱わない。浮動小数点の許容値による監査であり、区間演算の厳密数学証明ではない。P06N等のNケースはDavis等価関連モデルの離散化精度である。',
           '浸透ケースは乾燥モデルへ再構築しない。結合モデルの完全snapshotがないケースは今回の達成として数えない。',
           '',f'旧表で未達だった {len(previously_unmet)} ケースのうち、今回再確認した1%達成は {sum(c["target_met"] for c in previously_unmet)} ケース。',
           f'全26行では今回の再確認が {sum(c["target_met"] for c in case_reports)} / {len(case_reports)} ケース。旧表の達成と今回の再確認は区別する。','',
           '| ケース | 旧幅 % | 新下界Fs | 新上界Fs | 新幅 % | 両根完了 | 今回1%達成 | 状態 |',
           '|---|---:|---:|---:|---:|:---:|:---:|---|']
    for row in comparison:
        lines.append(f'| {row["case"]} | {fmt(row["old_gap_percent"],3)} | {fmt(row["new_lower_fs"])} | {fmt(row["new_upper_fs"])} | {fmt(row["new_gap_percent"],3)} | {"はい" if row["new_root_complete"] else "未確認"} | {"達成" if row["new_target_met"] else "未達/未再確認"} | {row["new_status"]} |')
    lines+=['','証拠場は snapshots/<case>/ に各側のmesh・field・certificate JSONを独立保存した。全候補の再監査と不合格・root未完了・更新中の不整合は rollup.json に記録した。',
            '比較CSVは COMPARISON_26.csv。旧速度測定と新しいPython実験の並行wall時間を混ぜた速度優位は主張しない。','',
            f'記録した不合格・skip件数: {len(issues)}。root未完了の物理許容場は端点証拠として保持するが、1%達成とはしない。','']
    (output/'REPORT_JA.md').write_text('\n'.join(lines),encoding='utf8')
    print('REPORT',output,'targets',sum(c['target_met'] for c in case_reports),'/',len(case_reports),'skipped',len(issues),flush=True)
    return payload


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--sources',nargs='+',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--table',type=Path,default=DEFAULT_TABLE)
    parser.add_argument('--tolerance',type=float,default=1e-4)
    args=parser.parse_args()
    if not 0<args.tolerance<=1e-4:raise ValueError('The final root tolerance cannot be relaxed')
    run(args.sources,args.output,args.table,args.tolerance)
