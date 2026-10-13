"""Compare the fast scaling with the prior formula on assembled programs."""
import sys
from pathlib import Path
import unittest

import numpy as np
from scipy import sparse

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import accuracy_driver as driver
import experiment_core as core


def prior_scales(matrix,dimensions):
    result=np.ones(matrix.shape[0]);start=0
    for dimension in dimensions:
        block=matrix[start:start+dimension]
        norm=float(np.sqrt(block.multiply(block).sum()))
        if norm>1e-10:result[start:start+dimension]=1/norm
        start+=dimension
    return result


class ConeScalingTests(unittest.TestCase):
    def test_same_positive_block_multiplier_on_real_lower_and_upper_programs(self):
        for case,fs in [('D01A',.58),('P06A',.85),('P06N',.68)]:
            mesh,_=driver.load_model(case)
            for mode in ['lower','upper']:
                with self.subTest(case=case,mode=mode):
                    prepared=core.Prepared(mesh,fs,mode,8,reduction='affine')
                    _,_,G,_,dimensions,_=prepared.reduced
                    fast=core.cone_row_scales(G,dimensions)
                    prior=prior_scales(G,dimensions)
                    np.testing.assert_allclose(fast,prior,rtol=4e-15,atol=0.)
                    self.assertTrue(np.all(np.isfinite(fast)&(fast>0)))
                    start=0
                    for dimension in dimensions:
                        self.assertTrue(np.all(fast[start:start+dimension]==fast[start]))
                        start+=dimension

    def test_zero_blocks_stay_unscaled_for_all_sparse_formats(self):
        matrix=sparse.coo_matrix(([3.,4.,1e-12],([0,1,4],[0,2,1])),shape=(6,3))
        for fmt in ['csc','csr','coo']:
            with self.subTest(format=fmt):
                np.testing.assert_array_equal(core.cone_row_scales(matrix.asformat(fmt),[2,2,1,1]),
                                              [.2,.2,1.,1.,1.,1.])
        np.testing.assert_array_equal(core.cone_row_scales(sparse.csc_matrix((0,3)),[]),[])

    def test_incomplete_or_nonpositive_dimensions_are_rejected(self):
        for dimensions in [[2],[1,0,2],[1,-1,3],[[1,2]]]:
            with self.subTest(dimensions=dimensions),self.assertRaises(ValueError):
                core.cone_row_scales(sparse.csc_matrix((3,2)),dimensions)


if __name__=='__main__':unittest.main()
