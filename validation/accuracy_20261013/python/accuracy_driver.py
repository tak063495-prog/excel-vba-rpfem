"""Audited adaptive Fs experiments; no certificate is inferred from a solver status.

Run with Python -X utf8. Bounds carry their own mesh and field. Failed trials
never overwrite a witness, and numerical signs only bracket the discrete root.
"""
from __future__ import annotations
import argparse
from dataclasses import dataclass
from pathlib import Path
import json
import platform
import struct
import time
import experiment_core as core
from experiment_core import np, Prepared, LowerRefresh, evaluate, solve, write

REPO = Path(__file__).resolve().parents[3]
CATALOG = REPO / 'docs/benchmarks/G1R17_native26_20261013/data/inputs'
GEOMETRY = REPO / 'versions/G1R17/results/geometry25'
ACCURATE = {'equilibrate_enable': False}


def load_model(case):
    rec = json.loads((CATALOG / f'{case}.json').read_text(encoding='utf8'))
    if rec['source_case'].get('pore'):
        raise ValueError('Pore pressure needs a coupled snapshot; dry reconstruction is prohibited')
    raw = (GEOMETRY / f'{case}_new_vba_mesh.bin').read_bytes()
    nn, ne = struct.unpack_from('<ii', raw)
    xy = np.frombuffer(raw, '<f8', 2 * nn, 8).copy().reshape(nn, 2)
    cells = np.frombuffer(raw, '<i4', 4 * ne, 8 + 16 * nn).copy().reshape(ne, 4)
    return core.mesh_from_arrays(xy, cells[:, :3], cells[:, 3], rec), rec


def load_saved_mesh(path, rec):
    with np.load(path) as data:
        return core.mesh_from_arrays(data['xy'].copy(), data['tri'].copy(), data['material_id'].copy(), rec)


def save_mesh(path, mesh):
    np.savez_compressed(path, xy=mesh.xy, tri=mesh.tri, material_id=mesh.material_id)


@dataclass
class Certificate:
    mode: str
    fs: float
    mesh: object
    field: object
    result: dict
    source: str
    root_complete: bool = False

    def publish(self, folder):
        stem = folder / f'best_{self.mode}'
        save_mesh(stem.with_name(stem.name + '_mesh.npz'), self.mesh)
        np.savez_compressed(stem.with_name(stem.name + '_field.npz'), x=self.field, fs=self.fs, mode=self.mode, q=8)
        return {'case':self.mesh.provenance['case'],'mode':self.mode,'q':8,'fs': self.fs, 'mesh_id': self.mesh.mesh_id, 'physical_model_id': self.mesh.physical_model_id,
                'elements': len(self.mesh.tri), 'source': self.source, 'root_complete': self.root_complete,
                'mesh_path': stem.name + '_mesh.npz', 'field_path': stem.name + '_field.npz', 'audit': self.result}


def finite(*values):
    try:
        return bool(np.isfinite(np.asarray(values,dtype=float)).all())
    except (TypeError,ValueError):
        return False


def sign(result, mode):
    target = core.LOWER_TARGET if mode == 'lower' else core.UPPER_TARGET
    if result.get('witness',False) and result.get('physical_pass',False) and finite(result.get('value')):
        return 1 if mode == 'lower' else -1
    if result.get('numerical_gate',False) and finite(result.get('value')):
        return 1 if result['value'] >= target else -1
    if not finite(result.get('dual_stationarity'),result.get('dual_cone_error'),result.get('dual_objective')):
        return None
    if result.get('dual_stationarity', 1) > 1e-7 or result.get('dual_cone_error', 1) > 1e-8:
        return None
    if mode == 'lower' and -result['dual_objective'] <= target - 2e-6:
        return -1
    if mode == 'upper' and result['dual_objective'] >= target + 2e-6:
        return 1
    return None


def search(mesh, mode, initial, folder, label, tolerance=1e-4, seconds=120, initial_witness=None):
    folder.mkdir(parents=True, exist_ok=True)
    refresh = LowerRefresh(mesh, 'affine') if mode == 'lower' else None
    if mode=='upper' and np.all(mesh.materials[:,1:3]==0) and not np.any(mesh.body[:,:,0]):
        from zero_phi_transport import PhiZeroUpperRefresh
        refresh=PhiZeroUpperRefresh(mesh,fs=1.,q=8,reduction='affine')
    target = core.LOWER_TARGET if mode == 'lower' else core.UPPER_TARGET
    cache, trials = {}, []
    best = None
    started = time.perf_counter()
    if initial_witness is not None:
        if initial_witness.mode!=mode or initial_witness.mesh.mesh_id!=mesh.mesh_id or not initial_witness.result['witness']:
            raise ValueError('An initial witness must own this exact mesh and side')
        result=dict(initial_witness.result)
        result.pop('discrete_root_search',None)
        result['origin']='independently_reaudited_same_mesh_witness'
        cache[initial_witness.fs]=(result,initial_witness.field.copy())
        trials.append(result)
        best=Certificate(mode,initial_witness.fs,mesh,initial_witness.field.copy(),result,label)

    def point(fs):
        nonlocal best
        fs = float(fs)
        if fs in cache:
            return cache[fs]
        prepared = refresh.at(fs) if refresh else Prepared(mesh, fs, mode, 8, reduction='affine')
        attempts = []
        dual_evidence = []
        chosen = None
        # Choose physical evidence before numerical optimality. A later failed
        # recovery cannot replace an earlier, independently audited witness.
        for options,scaled in [(ACCURATE,False), (dict(ACCURATE, iterative_refinement_max_iter=30,
                                      iterative_refinement_abstol=1e-15, iterative_refinement_reltol=1e-15),
                                      False),
                        (dict(ACCURATE, tol_gap_abs=1e-13, tol_gap_rel=1e-13, tol_feas=1e-13,
                              iterative_refinement_max_iter=30,iterative_refinement_abstol=1e-15,
                              iterative_refinement_reltol=1e-15),True),
                        ({'equilibrate_enable': True},False)]:
            result, field, _ = solve(prepared, dict(options, time_limit=seconds), max_iter=350,cone_scaling=scaled,primal_polish=True)
            attempts.append({k: result[k] for k in ['status', 'physical_pass', 'witness', 'numerical_gate', 'end_to_end_seconds']})
            dual_evidence.append({k: result.get(k) for k in ['dual_objective','dual_stationarity','dual_cone_error',
                                    'reduced_dual_objective','reduced_dual_stationarity','reduced_dual_cone_error']})
            rank = (result['witness'], result['physical_pass'], sign(result, mode) is not None, result['numerical_gate'])
            if chosen is None or rank > chosen[0]:
                chosen = rank, result, field
            if sign(result,mode) is not None:
                break
        _, result, field = chosen
        # Separate iterates at the same Fs may provide the admissible primal
        # and the admissible dual. Their gap need not be small for a dual to
        # screen a discrete root; the published bound is still only the primal.
        if sign(result,mode) is None:
            for evidence in dual_evidence:
                merged=dict(result,**evidence)
                if sign(merged,mode) is not None:
                    result=merged
                    result['dual_from_recovery_at_same_fs']=True
                    break
        if mode=='lower' and not result.get('witness') and sign(result,mode)!=-1 and result.get('value',0)>=.98:
            # A feasibility solve avoids maximizing a nearly singular load
            # variable. Its failure is UNKNOWN; success is only a primal bound.
            feasibility=Prepared(mesh,fs,'lower',8,reduction='affine',fixed_load=1.000003)
            options=dict(ACCURATE,iterative_refinement_max_iter=30,iterative_refinement_abstol=1e-15,
                         iterative_refinement_reltol=1e-15,time_limit=seconds)
            witness,wx,_=solve(feasibility,options,max_iter=350,cone_scaling=True,primal_polish=True)
            attempts.append({'fixed_load':1.000003,**{k:witness[k] for k in ['status','physical_pass','witness','end_to_end_seconds']}})
            if witness['witness']:
                # The fixed-load dual is for a different objective and must
                # never screen the load-maximization root.
                witness['numerical_gate']=False
                witness['reduced_numerical_gate']=False
                for key in list(witness):
                    if 'dual' in key:del witness[key]
                witness['fixed_load_feasibility_only']=True
                result,field=witness,wx
        result.update(attempts=attempts, search_sign=sign(result, mode), fs=fs)
        cache[fs] = result, field
        trials.append(result)
        if result['witness'] and (best is None or (fs > best.fs if mode == 'lower' else fs < best.fs)):
            best = Certificate(mode, fs, mesh, field.copy(), result, label)
            best.publish(folder)
        write(folder / (label + '_trials.json'), trials)
        print('POINT', label, fs, result['value'], result['status'], result['physical_pass'], result['witness'], flush=True)
        return result, field

    lo, hi = map(float, initial)
    if not finite(lo,hi) or not 0<lo<hi:
        raise ValueError('An initial bracket must have finite positive, ordered endpoints')
    lr, _ = point(lo)
    hr, _ = point(hi)
    if sign(lr,mode) is None or sign(hr,mode) is None:
        # Distant initial strengths can fail numerically even on a nested
        # mesh. Probe near the preceding endpoint before abandoning it.
        middle=(lo+hi)/2
        for fs in [middle,middle-2e-5,middle+2e-5]:
            result,_=point(fs)
            if sign(result,mode)==1 and (sign(lr,mode)!=1 or fs>lo):
                lo,lr=fs,result
            elif sign(result,mode)==-1 and (sign(hr,mode)!=-1 or fs<hi):
                hi,hr=fs,result
            if sign(lr,mode)==1 and sign(hr,mode)==-1 and lo<hi:
                break
    # Expand just the missing side, keeping every physical witness obtained.
    for _ in range(5):
        if sign(lr, mode) == 1 and sign(hr, mode) == -1:
            break
        span = hi - lo
        if sign(lr, mode) == -1:
            lo = max(1e-5, lo - span)
            lr, _ = point(lo)
        elif sign(hr, mode) == 1:
            hi += span
            hr, _ = point(hi)
        else:
            break
    complete = False
    if initial_witness is not None and lo<initial_witness.fs<hi:
        warm_result,_=cache[initial_witness.fs]
        if mode=='lower':lo,lr=initial_witness.fs,warm_result
        else:hi,hr=initial_witness.fs,warm_result
    if sign(lr, mode) == 1 and sign(hr, mode) == -1 and lo<hi:
        for _ in range(32):
            if hi - lo <= tolerance * max(1., (hi + lo) / 2):
                complete = True
                break
            denominator = hr['value'] - lr['value']
            fs = (lo + hi) / 2 if denominator == 0 or not finite(denominator,lr['value'],hr['value']) else lo + (target - lr['value']) * (hi - lo) / denominator
            fs = float(np.clip(fs, lo + .2 * (hi - lo), hi - .2 * (hi - lo)))
            result, _ = point(fs)
            if sign(result, mode) is None:
                result, _ = point((lo + hi) / 2)
                fs = (lo + hi) / 2
            s = sign(result, mode)
            if s is None:
                break
            if s == 1:
                lo, lr = fs, result
            else:
                hi, hr = fs, result
    endpoint = lo if mode == 'lower' else hi
    if best is not None:
        best.root_complete = bool(complete and best.fs == endpoint)
    summary = {'mode': mode, 'root_complete': bool(best and best.root_complete), 'bracket': [lo, hi],
               'bracket_width': hi - lo, 'fs_tolerance': tolerance, 'endpoint': best.fs if best else None,
               'trials': len(trials), 'seconds': time.perf_counter() - started,
               'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id,'elements':len(mesh.tri),
               'q':8,'case':mesh.provenance['case'],'policy':mesh.provenance['psi_policy'],
               'bracket_signs':[sign(lr,mode),sign(hr,mode)],
               'opposite_numerical_endpoint_is_continuum_bound': False}
    write(folder / (label + '_summary.json'), summary)
    if best is not None:
        best.result['discrete_root_search'] = summary
        best.publish(folder)
    return best, summary


def run(args):
    mesh, rec = load_model(args.case)
    folder = args.output.resolve() / args.case
    folder.mkdir(parents=True, exist_ok=True)
    best, history = {}, []
    velocity = None
    if args.resume and (folder/'summary.json').exists():
        args.seed=folder
    seed_meta = {}
    if args.seed:
        if (args.seed/'summary.json').exists():
            seed_meta=json.loads((args.seed/'summary.json').read_text(encoding='utf8'))
            history=seed_meta.get('history', [])
        elif (args.seed/'import_audit.json').exists():
            seed_meta=json.loads((args.seed/'import_audit.json').read_text(encoding='utf8'))['certificates']
        mesh = load_saved_mesh(args.seed / 'best_upper_mesh.npz', rec)
        for mode in ['lower', 'upper']:
            m = load_saved_mesh(args.seed / f'best_{mode}_mesh.npz', rec)
            with np.load(args.seed / f'best_{mode}_field.npz') as d:
                x, fs = d['x'].copy(), float(d['fs'])
                if str(d['mode'])!=mode or int(d['q'])!=8:
                    raise ValueError('Seed mode or quadrature differs from this run')
            result = evaluate(Prepared(m, fs, mode, 8, reduction='none'), x)
            if not result['witness']:
                raise ValueError(f'{mode} seed failed independent physical audit')
            metadata=seed_meta.get(mode,{})
            original_result=metadata.get('audit', {})
            root=original_result.get('discrete_root_search', {})
            source_path=Path(metadata.get('source',''))
            if not root and (source_path/'import_audit.json').is_file():
                imported=json.loads((source_path/'import_audit.json').read_text(encoding='utf8'))['certificates'][mode]
                if imported['fs']==fs and imported['mesh_id']==m.mesh_id:
                    metadata=imported
                    root=metadata['audit'].get('discrete_root_search',{})
            bracket=root.get('bracket',[])
            ordered=bool(len(bracket)==2 and finite(*bracket) and 0<bracket[0]<bracket[1])
            narrow=bool(ordered and bracket[1]-bracket[0]<=args.tolerance*max(1.,sum(bracket)/2))
            identity=bool(metadata.get('fs')==fs and metadata.get('mesh_id')==m.mesh_id and
                metadata.get('physical_model_id')==m.physical_model_id and metadata.get('case')==args.case and
                metadata.get('mode')==mode and metadata.get('q')==8 and root.get('mesh_id')==m.mesh_id and
                root.get('physical_model_id')==m.physical_model_id and root.get('case')==args.case and
                root.get('policy')==m.provenance['psi_policy'] and root.get('q')==8 and root.get('mode')==mode)
            complete=bool(identity and narrow and root.get('bracket_signs')==[1,-1] and
                metadata.get('root_complete') and root.get('root_complete') and root.get('endpoint')==fs and
                fs==(bracket[0] if mode=='lower' else bracket[1]))
            if root:
                result['discrete_root_search']=root
            best[mode] = Certificate(mode, fs, m, x, result, str(args.seed), complete)
        velocity = best['upper'].field
    initial = (.4, 1.3)
    offset=max((r['cycle'] for r in history),default=0)+int(args.recheck)
    for step in range(args.cycles + 1):
        cycle=offset+step
        if step or velocity is None or args.recheck:
            if velocity is not None and (step or not args.recheck):
                parent=mesh
                mesh, refinement = core.local_refine(mesh, rec, core.indicator(mesh, velocity), args.theta)
                velocity=core.prolong(parent,mesh,velocity,'upper')
                if len(mesh.tri) > args.max_elements:
                    print('ELEMENT_CAP', len(mesh.tri), flush=True)
                    break
            else:
                refinement = None
            label = f'{args.case}_C{cycle}'
            save_mesh(folder / (label + '_mesh.npz'), mesh)
            row = {'cycle': cycle, 'elements': len(mesh.tri), 'refinement': refinement}
            for mode in ['lower', 'upper']:
                solve_mesh=best[mode].mesh if args.recheck and step==0 and mode in best else mesh
                if args.fan_lower and mode=='lower' and not (args.recheck and step==0):
                    from mesh_accuracy import angular_fan_refine
                    solve_mesh,fan_info=angular_fan_refine(mesh,rec,scope='star')
                    if len(solve_mesh.tri)>args.max_elements:
                        solve_mesh=mesh
                        fan_info['skipped']='element_cap'
                    row['lower_angular_fan']=fan_info
                    save_mesh(folder/(label+'_lower_mesh.npz'),solve_mesh)
                if mode in best:
                    old = best[mode].fs
                    delta = max(.002,.002*old) if args.recheck and step==0 else max(.02, .04 * old)
                    bracket = (max(1e-5, old - delta), old + delta)
                else:
                    bracket = initial
                warm=best.get(mode)
                if warm is not None and warm.mesh.mesh_id!=solve_mesh.mesh_id:warm=None
                candidate, summary = search(solve_mesh, mode, bracket, folder / label / mode, label + '_' + mode, args.tolerance, args.solve_seconds,warm)
                row[mode] = summary
                if mode == 'upper' and candidate is not None:
                    velocity = candidate.field
                if candidate is not None and (mode not in best or (candidate.fs > best[mode].fs if mode == 'lower' else candidate.fs < best[mode].fs) or
                      (candidate.fs==best[mode].fs and candidate.root_complete and not best[mode].root_complete)):
                    if mode in best and candidate.mesh.physical_model_id != best[mode].mesh.physical_model_id:
                        raise ValueError('Physical models differ')
                    best[mode] = candidate
            history.append(row)
        if set(best) == {'lower', 'upper'}:
            lb, ub = best['lower'].fs, best['upper'].fs
            if not 0 < lb <= ub:
                raise ValueError('Reversed or nonpositive bounds')
            width = 2 * (ub - lb) / (lb + ub)
            report = {'case': args.case, 'lower': best['lower'].publish(folder), 'upper': best['upper'].publish(folder),
                      'relative_width_upper': (ub - lb) / ub, 'relative_width_symmetric': width,
                      'one_percent_width_pass': width <= .01, 'root_complete': all(c.root_complete for c in best.values()),
                      'target_met': width <= .01 and all(c.root_complete for c in best.values()),
                      'floating_tolerance_certificate': True, 'strict_interval_arithmetic_proof': False,
                      'environment': {'python': platform.python_version(), 'platform': platform.platform(),
                                      'numpy': np.__version__, 'threads': 1}, 'history': history}
            write(folder / 'summary.json', report)
            print('BOUNDS', args.case, len(mesh.tri), lb, ub, width, flush=True)
            if report['target_met']:
                break
        if velocity is None or len(velocity) < 12 * len(mesh.tri):
            break


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--case', required=True)
    parser.add_argument('--seed', type=Path)
    parser.add_argument('--resume', action='store_true')
    parser.add_argument('--recheck', action='store_true',help='Re-solve each retained side on its own witness mesh first')
    parser.add_argument('--fan-lower',action='store_true',help='One angle-space split on each proposed lower mesh; retain independent witnesses')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--cycles', type=int, default=8)
    parser.add_argument('--max-elements', type=int, default=12000)
    parser.add_argument('--theta', type=float, default=.6)
    parser.add_argument('--tolerance', type=float, default=1e-4)
    parser.add_argument('--solve-seconds', type=float, default=120)
    run(parser.parse_args())
