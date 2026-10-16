"""Same upper field/scalar backoff candidates; every point re-audited in its original program."""
from pathlib import Path
import time
import experiment_core as core
import accuracy_driver as driver
np=core.np


def main():
    root=Path(__file__).resolve().parents[1];case='D02A'
    source=root/'results/d02_final'/case
    _,rec=driver.load_model(case);mesh=driver.load_saved_mesh(source/'best_upper_mesh.npz',rec)
    with np.load(source/'best_upper_field.npz') as data:
        x=data['x'].copy();old=float(data['fs'])
        assert str(data['mode'])=='upper' and int(data['q'])==8
    output=root/'results'/('velocity_continuation_'+str(time.time_ns()))
    output.mkdir();report={'case':case,'source_fs':old,'q':8,'mesh_id':mesh.mesh_id,'runs':[]}
    for fraction in (0.,1e-4,2.5e-4,5e-4,7.5e-4,1e-3):
        fs=old*(1-fraction);pr=core.Prepared(mesh,fs,'upper',8,reduction='none')
        assert np.array_equal(pr.original[-1],np.zeros_like(pr.original[-1]))
        for kind,multiplier in [('same_velocity',1.),('Fs_inverse_scalar',old/fs)]:
            field=multiplier*x;audit=core.evaluate(pr,field,None)
            report['runs'].append({'fractional_fs_decrease':fraction,'fs':fs,'kind':kind,
                'field_multiplier':multiplier,'audit':audit})
            if audit['witness'] and fs<old:
                name=f'fraction_{fraction}_{kind}_field.npz'
                np.savez_compressed(output/name,x=field,fs=fs,mode='upper',q=8)
                driver.save_mesh(output/name.replace('_field','_mesh'),mesh)
            core.write(output/'audit.json',report)
            print(fraction,kind,audit['physical_pass'],audit['witness'],audit['eq_error'],
                audit['cone_error'],flush=True)
    print('DIAGNOSTIC',output,flush=True)


if __name__=='__main__':main()
