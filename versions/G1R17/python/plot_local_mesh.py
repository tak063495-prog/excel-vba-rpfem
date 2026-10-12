"""Before/after scientific mesh figure from actual native exports."""
import struct,numpy as np,matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from support import ROOT
def read(path):
 raw=path.read_bytes();nn,ne=struct.unpack_from('<ii',raw)
 xy=np.frombuffer(raw,'<f8',nn*2,8).reshape(nn,2);t=np.frombuffer(raw,'<i4',ne*4,8+16*nn).reshape(ne,4)
 return xy,t
def main():
 fig,axs=plt.subplots(1,2,figsize=(10.5,4.5),constrained_layout=True)
 for ax,label,title in zip(axs,['old','new'],['Before: 243 elements\nLower load factor ~ 0 at Fs=0.5','G1R17: 245 elements\nLower load factor = 1.60649 at Fs=0.5']):
  xy,t=read(ROOT/'results/geometry25'/('P06N_'+label+'_vba_mesh.bin'))
  colors=np.array(['#dce6dd','#e0eafa','#eee5d8'])[t[:,3]]
  ax.add_collection(PolyCollection(xy[t[:,:3]],facecolors=colors,edgecolors='#718096',linewidth=.7))
  ax.scatter([23],[5],s=55,color='#bc3030',zorder=5);ax.annotate('Exposed zero-stress vertex\n(23, 5)',(23,5),xytext=(21.8,5.83),arrowprops={'arrowstyle':'->','color':'#bc3030'},fontsize=9,color='#bc3030')
  if label=='old':
   poly=xy[t[65,:3]];ax.plot(*np.vstack([poly,poly[0]]).T,color='#bc3030',lw=2);ax.text(22.4,5.07,'Element 66\nobtuse fan',fontsize=9,color='#bc3030')
  else:
   p=xy[-1];ax.scatter(*p,s=45,color='#16836a',zorder=6);ax.annotate('One interior node\nthree child triangles',p,xytext=(23.15,4.13),arrowprops={'arrowstyle':'->','color':'#16836a'},fontsize=9,color='#16836a')
  ax.set_xlim(21.6,25.3);ax.set_ylim(3.8,6.05);ax.set_aspect('equal');ax.set_title(title,fontsize=11);ax.set_xlabel('x [m]');ax.set_ylabel('y [m]');ax.grid(alpha=.12)
 (ROOT/'figures').mkdir(exist_ok=True);fig.savefig(ROOT/'figures/P06N_local_mesh.png',dpi=180);fig.savefig(ROOT/'figures/P06N_local_mesh.pdf');plt.close(fig)
if __name__=='__main__':main()
