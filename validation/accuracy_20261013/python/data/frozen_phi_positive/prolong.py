"""Unmodified prolong() from 01_Initial_Experiments/experiment_core.py.

Compare on the same host, since saved BLAS outputs can differ by CPU rounding.
"""
import numpy as np


def prolong(parent,child,x,mode):
    result=[];pids=child.provenance['parent_element_ids'];values=x[:18*len(parent.tri)].reshape(-1,6,3) if mode=='lower' else x.reshape(-1,6,2)
    for t,e in zip(child.tri,pids):
        vertices=child.xy[t];points=np.r_[vertices,(vertices+np.roll(vertices,-1,axis=0))/2]
        matrix=np.vstack([parent.xy[parent.tri[e]].T,np.ones(3)])
        bary=np.linalg.solve(matrix,np.vstack([points.T,np.ones(6)])).T
        a,b,c=bary.T
        if mode=='lower':
            B=np.column_stack([a*a,b*b,c*c,2*a*b,2*b*c,2*c*a]);v=B@values[e]
            v[3:]=2*v[3:]-.5*(v[:3]+np.roll(v[:3],-1,axis=0));result.append(v)
        else:
            N=np.column_stack([a*(2*a-1),b*(2*b-1),c*(2*c-1),4*a*b,4*b*c,4*c*a]);result.append(N@values[e])
    return np.r_[np.array(result).ravel(),x[-1]] if mode=='lower' else np.array(result).ravel()
