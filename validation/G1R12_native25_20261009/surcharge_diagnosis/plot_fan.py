from pathlib import Path
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

root=Path(__file__).resolve().parent
fig,axes=plt.subplots(1,2,figsize=(10,4),layout='constrained')
for ax,tag,title in zip(axes,['original','centroid_fan'],['Original: two-triangle fan','Candidate: centroid fans']):
    with np.load(root/(tag+'_mesh.npz')) as z:xy=z['xy'];tri=z['tri']
    node=int(np.flatnonzero(np.all(xy==[2.,10.],axis=1))[0]);fan=tri[np.any(tri==node,axis=1)]
    ax.triplot(xy[:,0],xy[:,1],fan,color='#28475e',linewidth=1.5)
    ax.plot(2,10,'o',color='#b7463c',markersize=7,zorder=4)
    for x in [2.15,2.4,2.65]:ax.annotate('',xy=(x,10.01),xytext=(x,10.2),arrowprops={'arrowstyle':'->','color':'#b7463c'})
    ax.text(1.55,10.13,'Free',ha='center',fontsize=10)
    ax.text(2.48,10.23,'20 kPa',ha='center',fontsize=10,color='#b7463c')
    ax.set(xlim=(1.2,2.9),ylim=(9.1,10.4),xlabel='x [m]',ylabel='y [m]',title=title,aspect='equal')
    ax.grid(alpha=.18)
    ax.text(.02,.02,'Load 1 excluded by linear equations' if tag=='original' else 'Fs=1 / load=1 physical witness audited',transform=ax.transAxes,fontsize=9)
fig.savefig(root/'fan_repair.png',dpi=170)
fig.savefig(root/'fan_repair.svg')
