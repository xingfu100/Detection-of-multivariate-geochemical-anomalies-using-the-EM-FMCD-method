"""Draw figures for the synthetic case using the manuscript layout."""
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.ticker import AutoMinorLocator, MaxNLocator, MultipleLocator

ROOT=Path(__file__).resolve().parents[1]
out=ROOT/'results'
r=pd.read_csv(out/'sample_comparison.csv')
load=pd.read_csv(out/'rpca_loadings.csv')
summary=pd.read_csv(out/'model_summary.csv')
plt.rcParams.update({'font.family':'serif','font.size':9,'axes.linewidth':.8,'xtick.direction':'in','ytick.direction':'in','xtick.top':True,'ytick.right':True})
colors=['#154a99','#e60012','#ffed00','#68b72e','#70c8d6']

fig,ax=plt.subplots(figsize=(5.0,5.0),constrained_layout=True)
for i in range(5):
    sub=r[r.cluster==i]
    ax.scatter(sub.X_coord,sub.Y_coord,s=22,marker='s',c=colors[i],edgecolors='white',linewidths=.14,label=f'Cluster {i+1}')
ax.set(xlim=(-1,40),ylim=(-1,30),aspect='equal',xlabel='X (grid index)',ylabel='Y (grid index)')
ax.xaxis.set_major_locator(MultipleLocator(10));ax.yaxis.set_major_locator(MultipleLocator(10));ax.xaxis.set_minor_locator(AutoMinorLocator(2));ax.yaxis.set_minor_locator(AutoMinorLocator(2));ax.legend(loc='lower center',bbox_to_anchor=(.5,-.32),ncol=3,frameon=False,fontsize=8)
fig.savefig(out/'figure_em.png',dpi=300,bbox_inches='tight');plt.close(fig)

fig,axes=plt.subplots(2,3,figsize=(10,7),constrained_layout=True)
colors_anom=['#ffed00','#f58b00','#d40019']
for i,ax in enumerate(axes.flat):
    sub=r if i==5 else r[r.cluster==i]; col='regional_level' if i==5 else 'clustered_level'
    for level,color in zip([1,2,3],colors_anom):
        points=sub[sub[col]==level];ax.scatter(points.X_coord,points.Y_coord,s=12,c=color,marker='s',linewidths=0)
    ax.set(xlim=(-1,40),ylim=(-1,30),aspect='equal',xlabel='X (grid index)',ylabel='Y (grid index)')
    ax.text(.03,.96,f'({chr(97+i)})',transform=ax.transAxes,ha='left',va='top',fontweight='bold');ax.xaxis.set_major_locator(MultipleLocator(10));ax.yaxis.set_major_locator(MultipleLocator(10));ax.xaxis.set_minor_locator(AutoMinorLocator(2));ax.yaxis.set_minor_locator(AutoMinorLocator(2))
fig.legend(handles=[Line2D([0],[0],marker='s',linestyle='',color=c,label=l) for c,l in zip(colors_anom,['Weak anomaly','Moderate anomaly','Strong anomaly'])],loc='outside lower center',ncol=3,frameon=False)
fig.savefig(out/'figure_anomalies.png',dpi=300);plt.close(fig)

fig,axes=plt.subplots(3,2,figsize=(7.5,10),constrained_layout=True)
for i,ax in enumerate(axes.flat):
    key=str(i) if i<5 else 'regional'; v=load[load.model.astype(str)==key].copy(); v['magnitude']=np.hypot(v.PC1,v.PC2)
    # Limit labels to the largest vectors and the selected association.
    pb_length=float(v.loc[v.element.eq('Pb'),'magnitude'].iloc[0])
    v=v[(v.magnitude >= pb_length/2) & (v.magnitude <= pb_length*2)]
    v=v[v.selected | v.index.isin(v.nlargest(6,'magnitude').index)]
    ax.axhline(0,c='#bfd0d8',lw=.5);ax.axvline(0,c='#bfd0d8',lw=.5)
    extent=max(float(v.PC1.abs().max()), float(v.PC2.abs().max())) * 1.8
    extent=max(float(v.PC1.abs().max()), float(v.PC2.abs().max())) * 1.55
    placed=[]
    for _,row in v.iterrows():
        ax.annotate('',xy=(row.PC1,row.PC2),xytext=(0,0),arrowprops={'arrowstyle':'-|>','lw':.9,'color':'black','mutation_scale':6})
        if row.selected:
            label_y=float(row.PC2)
            for _ in range(8):
                if not any(abs(row.PC1-px)<.10*extent and abs(label_y-py)<.08*extent for px,py in placed):break
                label_y+=.08*extent
            label_y=min(label_y,.90*extent)
            placed.append((float(row.PC1),label_y))
            ax.annotate(row.element,xy=(row.PC1,row.PC2),xytext=(row.PC1,label_y),fontsize=8,clip_on=True)
    ax.set_xlim(-extent, extent); ax.set_ylim(-extent, extent)
    sr=summary[summary.model.astype(str)==key].iloc[0]
    ax.set_xlabel(f"PC1 ({100*sr.pc1_fraction:.2f}%)");ax.set_ylabel(f"PC2 ({100*sr.pc2_fraction:.2f}%)")
    ax.text(.04,.96,f'({chr(97+i)})',transform=ax.transAxes,ha='left',va='top',fontweight='bold');ax.set_aspect('equal',adjustable='box');ax.xaxis.set_major_locator(MaxNLocator(nbins=4));ax.yaxis.set_major_locator(MaxNLocator(nbins=4));ax.xaxis.set_minor_locator(AutoMinorLocator(2));ax.yaxis.set_minor_locator(AutoMinorLocator(2))
fig.savefig(out/'figure_rpca.png',dpi=300);plt.close(fig)
print('Saved EM map, RPCA loadings, and clustered/regional anomaly maps.')



