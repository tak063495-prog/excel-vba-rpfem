import numpy as np

P2_BARY=np.array([[1,0,0],[0,1,0],[0,0,1],[.5,.5,0],[0,.5,.5],[.5,0,.5]],float)

def shape(lam):
    l=np.asarray(lam,float)
    return np.concatenate((l*(2*l-1),4*l*np.roll(l,-1,axis=-1)),axis=-1)

def gradients(lam, grad):
    l=np.asarray(lam,float)
    return np.concatenate(((4*l-1)[...,None]*grad,
                           4*(l[...,None]*np.roll(grad,-1,axis=0)+np.roll(l,-1,axis=-1)[...,None]*grad)),axis=-2)

def trace(s):
    return np.array([(1-s)*(1-2*s),s*(2*s-1),4*s*(1-s)])

def bernstein(a,b):
    left,right=trace(a),trace(b)
    return np.array([left,2*trace((a+b)/2)-(left+right)/2,right])

def contained_intervals(start,end,old_q,new_q):
    """Actual containment test; does not assume arbitrary splits are nested."""
    points=np.linspace(start,end,new_q+1)
    return all(any(a>=j/old_q-1e-14 and b<=(j+1)/old_q+1e-14 for j in range(old_q))for a,b in zip(points[:-1],points[1:]))
