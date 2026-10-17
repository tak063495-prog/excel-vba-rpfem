"""Audited exact strength-scaling transport for dry, reference-only phi=0.

This produces a new one-sided physical certificate. It makes no optimality
claim and does not mark a discrete root search complete.
"""
import time
import experiment_core as core
from experiment_core import np

def require_reference_phi_zero(mesh):
    if mesh.load_phase!='reference_total':
        raise ValueError('Phi=0 scaling requires the reference_total load phase')
    if not np.all(mesh.materials[:,1:3]==0) or not np.all(mesh.rigid_id<0) or any(b.get('rigid',-1)>=0 for b in mesh.boundary.values()):
        raise ValueError('Transport requires only soil with phi=psi=0')
    if np.any(mesh.body[:,:,0]!=0) or any(np.any(np.asarray(b['loads'])[:,0]!=0) for b in mesh.boundary.values()):
        raise ValueError('Transport requires zero fixed body and boundary loads')


class PhiZeroUpperRefresh:
    """Keep the normal root search, reusing its unchanged phi=0 E/G maps.

    Only the cohesive dissipation prices depend on Fs in this restricted
    model. Every candidate still uses the ordinary uncondensed physical audit.
    """
    def __init__(self,mesh,fs=1.,q=8,reduction='affine'):
        require_reference_phi_zero(mesh)
        self.base=core.Prepared(mesh,fs,'upper',q,reduction=reduction)
        self.mesh_id=mesh.mesh_id

    def at(self,fs):
        if not np.isfinite(fs) or fs<=0:raise ValueError('Fs must be finite and positive')
        if self.base.m.mesh_id!=self.mesh_id:raise ValueError('Cached mesh changed')
        started=time.perf_counter();old=self.base;ratio=old.fs/fs
        p=core.Prepared.__new__(core.Prepared);p.__dict__=old.__dict__.copy();p.fs=fs
        p.original=(*old.original[:-1],old.original[-1]*ratio)
        p.cost=old.cost*ratio;p.objective_offset=old.objective_offset*ratio
        p.reduced=(*old.reduced[:-1],p.cost)
        p.assembly_seconds=time.perf_counter()-started;p.transform_seconds=0.
        return p


def transport(mesh,fs,field,mode,q=8,target=None):
    started=time.perf_counter()
    if mode not in ('lower','upper'):
        raise ValueError('Only lower/upper certificates can be transported')
    require_reference_phi_zero(mesh)
    if target is None:target=1.0000012 if mode=='lower' else .9999988
    if not np.isfinite(target) or target<=0 or (target<core.LOWER_TARGET if mode=='lower' else target>core.UPPER_TARGET):
        raise ValueError('Transport target cannot weaken the physical witness margin')
    field=np.asarray(field,float)
    before=core.evaluate(core.Prepared(mesh,fs,mode,q,reduction='none'),field)
    source_seconds=time.perf_counter()-started
    if not before['physical_pass'] or before['value']<=0:
        raise ValueError('The source field must pass its original physical audit')
    new_fs=float(fs*before['value']/target)
    # Stress, body-load multiplier, and strength scale together in the lower
    # field. Upper velocities, normalization, and norm epigraphs stay fixed.
    alpha=target/before['value'] if mode=='lower' else 1.
    transported=field*alpha
    after=core.evaluate(core.Prepared(mesh,new_fs,mode,q,reduction='none'),transported)
    record={'mode':mode,'q':q,'source_fs':float(fs),'fs':new_fs,'source_value':before['value'],
            'value':after['value'],'target':float(target),'field_multiplier':float(alpha),
            'physical_pass':after['physical_pass'],'witness':after['witness'],
            'root_complete':False,'physical_model_id':mesh.physical_model_id,
            'mesh_id':mesh.mesh_id,'audit':after,
            'source_audit_seconds':source_seconds,
            'transport_and_audit_seconds':time.perf_counter()-started-source_seconds,
            'total_seconds':time.perf_counter()-started,
            'meaning':'AUDITED_STRENGTH_SCALING_PRIMAL_WITNESS_ONLY'}
    if not after['physical_pass'] or not after['witness']:
        raise ValueError(f'Transported field failed the unchanged physical gate: {record}')
    return record,transported
