"""Generate a fictional 40 x 30 geological-style synthetic dataset."""
from pathlib import Path
import numpy as np
import pandas as pd
from scipy.ndimage import gaussian_filter

COLUMNS = ['Be','Cr','F','La','Li','Nb','P','Ti','V','Y','Al2O3','Fe2O3','K2O','SiO2','Ag','As','Au','B','Ba','Bi','Cd','Co','Cu','Hg','Mn','Mo','Ni','Pb','Sb','Sn','Sr','Th','U','W','Zn','Zr','CaO','MgO','Na2O']
ASSOCIATIONS = [['Pb','Zn','Cu','F'], ['Pb','W','B','Bi'], ['Pb','Bi','Sn'], ['Pb','Cu','Zn','Ti','Y','Zr'], ['Pb','Li','Be','Sn','Y']]

def generate():
    rng=np.random.default_rng(42); yy,xx=np.mgrid[0:30,0:40]
    # Interleaved smoothed fields plus two distinct irregular bodies.
    fields=np.array([gaussian_filter(rng.normal(size=(30,40)),s)*s for s in (3.5,4.0,4.5)])
    fields += np.array([.10*np.sin((xx+yy)/5),.12*np.cos((2*xx-yy)/8),.10*np.sin((xx-2*yy)/7)])
    zone=np.argmax(fields,axis=0)
    warp=gaussian_filter(rng.normal(size=(30,40)),2)*4
    body1=((xx-28)/8.0)**2+((yy-9)/6.0)**2 < 1+warp
    body2=((xx-10)/5.5)**2+((yy-22)/6.0)**2 < 1+warp*.8
    zone[body1]=3;zone[body2]=4
    coords=np.column_stack((xx.ravel(),yy.ravel())); region=zone.ravel().copy(); n=1200
    # Keep the random stream used by the previous example while removing
    # isolated reassignment; spatial mixing remains confined to contacts.
    previous_mask = rng.random(n) < .07
    rng.integers(0, 5, previous_mask.sum())
    # Smooth background signatures intentionally differ by synthetic zone.
    signatures=rng.normal(0,.55,(5,14))
    # Moderately vary major oxides across backgrounds while retaining their scale.
    signatures[:, 10:14] *= .18
    # All synthetic concentrations are in a common ppm-like scale. Rounded
    # parameters approximate broad observed ranges, not individual records.
    location = np.full(39, 20.0)
    spread = np.full(39, .45)
    # Coarse simulated scales only; no observed sample values are used.
    levels={'Ag':20,'As':20,'Au':.5,'B':20,'Ba':300,'Be':3,'Bi':.5,'Cd':200,
            'Co':20,'Cr':50,'Cu':30,'F':300,'Hg':20,'La':50,'Li':30,'Mn':500,
            'Mo':.5,'Nb':20,'Ni':30,'P':400,'Pb':30,'Sb':3,'Sn':4,'Sr':50,
            'Th':20,'Ti':2000,'U':3,'V':50,'W':4,'Y':20,'Zn':50,'Zr':200}
    for name,center in levels.items():
        location[COLUMNS.index(name)] = center
    for name,center,sigma in [('Sn',4,1.0),('Al2O3',120000,.30),
                              ('CaO',5500,.95),('K2O',17000,.55),
                              ('MgO',6500,.40),('Na2O',1700,.90),
                              ('SiO2',680000,.12)]:
        location[COLUMNS.index(name)] = center
        spread[COLUMNS.index(name)] = sigma
    values=rng.normal(np.log(location),spread,(n,39))
    # Smooth mixtures model transitional signatures at contacts; every sample
    # also has individual noise, so components can overlap near boundaries.
    fractions=np.stack([gaussian_filter((zone==z).astype(float),1.4) for z in range(5)],axis=-1).reshape(n,5)
    weights=.75*np.eye(5)[region]+.25*fractions
    values[:,:14]+=weights@signatures
    for z in range(5):
        values[region==z,COLUMNS.index('Pb')] += [0,.15,.15,.35,1.1][z]
    # A fictional lithological background shift is shared by Pb and its
    # associated trace elements. These offsets are constant within the unit.
    # EM indicator concentrations are not directly changed; CLR values may change
    # because the CLR denominator includes all 39 elements.
    for element, shift in [('Pb', 1.0), ('Bi', .85), ('Sn', .65), ('W', .45)]:
        values[region==4, COLUMNS.index(element)] += shift
    # Artificial anomalies: irregular clustered enrichment along a bent belt.
    trend=np.exp(-((yy-(5+.38*xx+2*np.sin(xx/7)))/3.4)**2)
    spots=sum(np.exp(-((xx-x)**2+(yy-y)**2)/20) for x,y in [(6,13),(14,19),(25,19),(32,20),(28,9),(10,22)])
    strength=(.75*trend+.95*spots).ravel(); perturb=rng.normal(0,.04,n); mineral=1.5*np.maximum(0,strength+perturb)
    # Group-specific correlated fields give each composition its own
    # covariance structure and keep associations distinguishable.
    for z,association in enumerate(ASSOCIATIONS):
        mask=region==z
        latent1=rng.normal(0,.18,mask.sum()) + mineral[mask]*.9
        latent2=rng.normal(0,.25,mask.sum())
        for j,element in enumerate(association):
            values[mask,COLUMNS.index(element)] += latent1*(1.0 if element=='Pb' else .85)
        partners=[e for e in COLUMNS[14:] if e not in association]
        for element in partners[z*3:z*3+4]:
            values[mask,COLUMNS.index(element)] += latent2*.8
    data=np.exp(values)
    oxides=['Al2O3','CaO','K2O','MgO','Na2O','SiO2','Fe2O3']
    positions=[COLUMNS.index(name) for name in oxides]
    totals=data[:,positions].sum(axis=1)
    scale=np.minimum(1.,950000. / totals)
    data[:,positions] *= scale[:,None]
    df=pd.DataFrame(data,columns=COLUMNS);df.insert(0,'Y_coord',coords[:,1]);df.insert(0,'X_coord',coords[:,0])
    truth=pd.DataFrame({'X_coord':coords[:,0],'Y_coord':coords[:,1],'synthetic_zone':region,'synthetic_lithology':np.array(['carbonate','sandstone','shale','granite','mudstone'])[region],'imposed_intensity':strength,'imposed_anomaly':strength>=.55})
    # A single seeded random shuffle avoids spatial ordering in model input.
    order = np.arange(n)
    attempt = 0
    return df.iloc[order].reset_index(drop=True),truth.iloc[order].reset_index(drop=True),attempt

if __name__=='__main__':
    data,truth,attempt=generate(); folder=Path(__file__).resolve().parent
    data.to_csv(folder/'synthetic_1200x39.csv',index=False,float_format='%.8f')
    truth.to_csv(folder/'synthetic_ground_truth.csv',index=False,float_format='%.8f')
    print('Generated',len(data),'fictional samples; row-order attempt',attempt,'zone sizes',truth.synthetic_zone.value_counts().sort_index().to_dict())

