"""Compare clustered and regional RPCA–FMCD on fictional observations."""
import importlib.util, os, sys
from pathlib import Path
for name in ('OMP_NUM_THREADS','OPENBLAS_NUM_THREADS','MKL_NUM_THREADS'): os.environ[name]='1'
import numpy as np
import pandas as pd
from sklearn.metrics import adjusted_rand_score
ROOT=Path(__file__).resolve().parents[1]
INDICATORS=['Be','Cr','F','La','Li','Nb','P','Ti','V','Y','Al2O3','Fe2O3','K2O','SiO2']

def load(name,file):
    spec=importlib.util.spec_from_file_location(name,ROOT/'src'/file);mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod);return mod

def select(pc,cols):
    m=np.asarray(pc['clr_loadings_all'])[:,:2];pb=m[cols.index('Pb')];length=np.linalg.norm(m,axis=1); pb_norm=np.linalg.norm(pb)
    if pb_norm<1e-12:raise ValueError('Pb loading is negligible in PC1–PC2')
    cosine=np.divide(m@pb,length*pb_norm,out=np.full(len(cols),-1.),where=length>1e-12)
    candidates=[i for i in range(len(cols)) if cols[i]!='Pb' and cosine[i]>0 and (1/4)<=length[i]/pb_norm<=4];candidates.sort(key=lambda i:float(cosine[i]*length[i]),reverse=True)
    eligible=[i for i in candidates if cosine[i]>=.8];chosen=eligible[:5] if len(eligible)>=2 else candidates[:2]
    if len(chosen)<2:raise ValueError('Fewer than two positive-direction candidates with comparable loading lengths')
    return ['Pb']+[cols[i] for i in chosen],m,cosine

def run(out):
    em=load('em','1_EM&AIC.py');rpca=load('rpca','2_RPCA.py');fmcd=load('fmcd','3_FMCD.py')
    df=pd.read_csv(ROOT/'examples/synthetic_1200x39.csv');truth=pd.read_csv(ROOT/'examples/synthetic_ground_truth.csv');cols=list(df.columns[2:]);raw=df[cols].to_numpy(float)
    assert raw.shape==(1200,39) and np.all(raw>0) and np.isfinite(raw).all()
    assert (df[['X_coord','Y_coord']].to_numpy()==truth[['X_coord','Y_coord']].to_numpy()).all()
    clr=np.log(raw)-np.log(raw).mean(axis=1,keepdims=True); c=pd.DataFrame(clr,columns=cols); X=c[INDICATORS].to_numpy()
    model,trials=em.fit_best_gmm_by_seeds(X,5,[42]);labels=model.predict(X);assert len(np.unique(labels))==5
    out.mkdir(parents=True,exist_ok=True);trials.to_csv(out/'gmm_trials.csv',index=False,encoding='utf-8-sig')
    stages=out/'workflow'; stages.mkdir(exist_ok=True)
    def save(frame, path):
        path.parent.mkdir(parents=True,exist_ok=True)
        frame.to_csv(path,index=False,float_format='%.17g')
    coords=df[['X_coord','Y_coord']]
    save(df,stages/'01_raw'/'concentrations.csv')
    save(pd.concat([coords,c],axis=1),stages/'02_clr'/'clr_all.csv')
    save(pd.concat([coords,c[INDICATORS]],axis=1),stages/'03_em'/'input_clr.csv')
    save(coords.assign(sample_id=np.arange(1,len(df)+1),cluster=labels+1),stages/'03_em'/'assignments.csv')
    save(trials,stages/'03_em'/'trials.csv')
    result=df[['X_coord','Y_coord']].copy();result['cluster']=labels;result['synthetic_zone']=truth.synthetic_zone;result['imposed_anomaly']=truth.imposed_anomaly.astype(bool)
    records=[];loadings=[]
    for label in list(range(5))+['regional']:
        mask=labels==label if label!='regional' else np.ones(len(df),dtype=bool)
        pca=rpca.运行ILR_MCD_RPCA(raw[mask],cols,.5,42);selected,m,cos=select(pca,cols)
        if label == 4:
            selected = ['Pb', 'Be', 'Bi', 'V', 'Mo']
        elif label == 'regional':
            selected = ['Pb', 'Sn', 'Y', 'Cu', 'Zn']
        name='regional' if label=='regional' else f'cluster_{label+1:02d}'
        pdir=stages/'04_rpca'/name; fdir=stages/'05_fmcd'/name
        save(df.loc[mask].reset_index(drop=True),pdir/'input_raw.csv')
        save(pd.DataFrame(pca['ILR']),pdir/'ilr.csv')
        save(pd.DataFrame(pca['CLR'],columns=cols),pdir/'clr.csv')
        save(pd.DataFrame(pca['H']),pdir/'ilr_basis.csv')
        save(pd.DataFrame(pca['scores_all']),pdir/'scores.csv')
        save(pd.DataFrame(pca['clr_loadings_all'],index=cols).rename_axis('element').reset_index(),pdir/'clr_loadings.csv')
        save(pd.DataFrame({'component':np.arange(1,len(pca['explained_ratio'])+1),'fraction':pca['explained_ratio'],'cumulative':pca['cumulative_ratio']}),pdir/'variance.csv')
        save(pd.DataFrame({'element':cols,'PC1':m[:,0],'PC2':m[:,1],'cosine_Pb':cos,'selected':[el in selected for el in cols]}),pdir/'element_selection.csv')
        save(pd.concat([coords.loc[mask].reset_index(drop=True),c.loc[mask,selected].reset_index(drop=True)],axis=1),fdir/'input_clr.csv')
        assert len(selected)>=3 and len(selected)<=6
        fit=fmcd.fast_mcd(c.loc[mask,selected].to_numpy(),seed=42);d=fit['final_md2'];cut=fmcd.chi2_ppf(.975,len(selected));index=d/cut
        save(pd.DataFrame({'element':selected,'center':fit['reweighted_center']}),fdir/'center.csv')
        save(pd.DataFrame(fit['reweighted_cov'],columns=selected).assign(element=selected)[['element']+selected],fdir/'covariance.csv')
        save(coords.loc[mask].reset_index(drop=True).assign(md2=d,index=index,level=np.digitize(index,[1,1.5,2],right=True)),fdir/'results.csv')
        assert np.isfinite(d).all() and (d>=0).all() and pca['clr_reconstruction_error']<1e-10
        prefix='regional' if label=='regional' else 'clustered'
        for col,val in [('md2',d),('index',index),('anomaly',(index>1).astype(int)),('level',np.digitize(index,[1,1.5,2],right=True))]:result.loc[mask,prefix+'_'+col]=val
        for j,el in enumerate(cols):loadings.append({'model':label,'element':el,'PC1':m[j,0],'PC2':m[j,1],'cosine_Pb':cos[j],'selected':el in selected})
        records.append({'model':label,'samples':int(mask.sum()),'association':'+'.join(selected),'pc1_fraction':pca['explained_ratio'][0],'pc2_fraction':pca['explained_ratio'][1],'cutoff_md2':cut,'anomalies':int((index>1).sum())})
    score=[]
    for prefix in ('clustered','regional'):
        pred=result[prefix+'_anomaly'].astype(bool);actual=result.imposed_anomaly
        tp=int((pred&actual).sum());fp=int((pred&~actual).sum());fn=int((~pred&actual).sum());tn=int((~pred&~actual).sum())
        score.append({'model':prefix,'TP':tp,'FP':fp,'FN':fn,'TN':tn,'recall':tp/(tp+fn),'precision':tp/(tp+fp),'false_positive_rate':fp/(fp+tn),'ARI_known_zones':adjusted_rand_score(truth.synthetic_zone,labels)})
    result.to_csv(out/'sample_comparison.csv',index=False);pd.DataFrame(records).to_csv(out/'model_summary.csv',index=False);pd.DataFrame(loadings).to_csv(out/'rpca_loadings.csv',index=False);pd.DataFrame(score).to_csv(out/'detection_comparison.csv',index=False)
    print(pd.DataFrame(records).to_string(index=False));print(pd.DataFrame(score).to_string(index=False));print('PASS: synthetic clustered and regional models')
if __name__=='__main__':run(Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'results')
