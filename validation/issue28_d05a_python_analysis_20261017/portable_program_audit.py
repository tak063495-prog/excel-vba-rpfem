"""Strict structure plus explicit floating roundoff checks for saved LPs.

These are coefficient/replay identity checks, never physical acceptance
tolerances. Original feasibility/dual/sign gates are evaluated separately.
"""
from __future__ import annotations
import hashlib
import numpy as np
from scipy import sparse

FLOAT_EPS_FACTOR=32
EPS=float(np.finfo(np.float64).eps)
FLOAT_KEYS=('E_data','b','G_data','h','cost')
EXACT_KEYS=('E_shape','E_indptr','E_indices','G_shape','G_indptr','G_indices','dims')
ALL_KEYS=set(FLOAT_KEYS+EXACT_KEYS)

def program_digest(program):
    """Same raw-byte digest used by the frozen original producer."""
    h=hashlib.sha256()
    for value in program:
        arrays=(value.indptr,value.indices,value.data)if sparse.issparse(value)else(np.asarray(value),)
        for array in arrays:
            h.update(str(array.shape).encode('ascii'));h.update(str(array.dtype).encode('ascii'));h.update(array.tobytes())
    return h.hexdigest()

def pack_program(program):
    E,b,G,h,dims,cost=program
    assert sparse.isspmatrix_csc(E)and sparse.isspmatrix_csc(G),'CSC_REQUIRED'
    return {'E_shape':np.asarray(E.shape,dtype=np.int64),'E_indptr':E.indptr.copy(),
        'E_indices':E.indices.copy(),'E_data':E.data.copy(),'b':np.asarray(b).copy(),
        'G_shape':np.asarray(G.shape,dtype=np.int64),'G_indptr':G.indptr.copy(),
        'G_indices':G.indices.copy(),'G_data':G.data.copy(),'h':np.asarray(h).copy(),
        'dims':np.asarray(dims).copy(),'cost':np.asarray(cost).copy()}

def unpack_program(arrays):
    assert set(arrays)==ALL_KEYS,'UNEXPECTED_REFERENCE_ARRAYS'
    for key in EXACT_KEYS:
        assert arrays[key].dtype.kind in 'iu','INTEGER_STRUCTURE_REQUIRED '+key
    for key in FLOAT_KEYS:
        assert arrays[key].dtype==np.dtype('float64'),'FLOAT64_REQUIRED '+key
        assert np.isfinite(arrays[key]).all(),'NONFINITE_REFERENCE '+key
    E=sparse.csc_matrix((arrays['E_data'],arrays['E_indices'],arrays['E_indptr']),shape=tuple(arrays['E_shape']))
    G=sparse.csc_matrix((arrays['G_data'],arrays['G_indices'],arrays['G_indptr']),shape=tuple(arrays['G_shape']))
    E.check_format(full_check=True);G.check_format(full_check=True)
    assert arrays['b'].shape==(E.shape[0],)and arrays['h'].shape==(G.shape[0],)
    assert E.shape[1]==G.shape[1]==len(arrays['cost'])
    assert arrays['dims'].ndim==1 and np.all(arrays['dims']>0)and int(arrays['dims'].sum())==G.shape[0]
    return E,arrays['b'],G,arrays['h'],arrays['dims'].tolist(),arrays['cost']

def float_roundoff_bound(reference):
    return FLOAT_EPS_FACTOR*EPS*max(1.,float(np.max(abs(reference),initial=0.)))

def compare_program(observed_program,reference_program):
    """Fail on layout/type/nonfinite/tamper; allow only 32eps-scaled floats.

    Sparse shape, CSC pointers/order/indices, integer cone dimensions and all
    dtypes are exact. Each floating array uses 32*eps*max(1,max(abs(reference)))
    as a coefficient comparison bound. A few-ULP edge-norm dispatch difference
    changes G/cost bits; the bound does not replace any physical or dual gate.
    """
    actual=pack_program(observed_program);expected=pack_program(reference_program);details=[]
    for key in EXACT_KEYS+FLOAT_KEYS:
        a,b=actual[key],expected[key]
        assert a.shape==b.shape,'SHAPE_MISMATCH '+key
        assert a.dtype==b.dtype,'DTYPE_MISMATCH '+key
        exact=bool(a.tobytes()==b.tobytes())
        if key in EXACT_KEYS:
            assert np.array_equal(a,b),'STRUCTURE_MISMATCH '+key
            details.append({'name':key,'kind':'exact_integer_structure','shape':a.shape,'dtype':str(a.dtype),'bitwise_equal':exact})
        else:
            assert np.isfinite(a).all()and np.isfinite(b).all(),'NONFINITE_COEFFICIENT '+key
            delta=float(np.max(abs(a-b),initial=0.));bound=float_roundoff_bound(b)
            assert delta<=bound,f'COEFFICIENT_ROUNDOFF_EXCEEDED {key} difference={delta:.17g} bound={bound:.17g}'
            details.append({'name':key,'kind':'float64_coefficient','shape':a.shape,'dtype':str(a.dtype),
                'bitwise_equal':exact,'different_entries':int(np.count_nonzero(a!=b)),
                'max_absolute_difference':delta,'roundoff_bound':bound,
                'max_abs_reference':float(np.max(abs(b),initial=0.))})
    return {'status':'PASS','raw_sha256_identical':program_digest(observed_program)==program_digest(reference_program),
        'observed_raw_sha256':program_digest(observed_program),'reference_raw_sha256':program_digest(reference_program),
        'float_comparison_formula':'32 * binary64_epsilon * max(1, max(abs(reference_array)))',
        'physical_acceptance_thresholds_changed':False,'arrays':details}

def compare_metric(observed,expected,label):
    """Narrow replay-value comparison; the acceptance booleans stay exact."""
    actual,reference=float(observed),float(expected)
    assert np.isfinite(actual)and np.isfinite(reference),'NONFINITE_AUDIT_METRIC '+label
    bound=FLOAT_EPS_FACTOR*EPS*max(1.,abs(reference))
    delta=abs(actual-reference)
    assert delta<=bound,f'AUDIT_METRIC_ROUNDOFF_EXCEEDED {label} difference={delta:.17g} bound={bound:.17g}'
    return {'name':label,'absolute_difference':delta,'roundoff_bound':bound,'bitwise_equal':np.float64(actual).tobytes()==np.float64(reference).tobytes()}
