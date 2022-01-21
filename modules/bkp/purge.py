# Purge Matrix extraction from Excel Eval Sheets
# by Lukas Brändli, 2021

import pandas as pd

def getPurge(sample1, sample2, data, inv=True):

#     sel =  [sample1, sample2]

    #   get purity data in accessible (non-PIVOT)-format & reset index
    if isinstance(data.index, pd.core.indexes.multi.MultiIndex):
        data_Purity = data[['Rel.Area']]
    
    data_Purity = data
    data_Purity = data_Purity.reset_index()
    
#     Extract sample data
    sample1_data = data_Purity.loc[data_Purity['Sample Type'] == sample1]
    sample2_data = data_Purity.loc[data_Purity['Sample Type'] == sample2]

# merge sample data: Match Experiments and Peak Names
    purge = pd.merge(sample1_data, sample2_data, on=['Experiment', 'Peak Name'], how='inner')
    # purge

    purge['purge'] = purge['Rel.Area_x'] / purge['Rel.Area_y']

    # purge = purge[['Experiment', 'Peak Name', 'purge']]
    purge = purge[['Experiment', 'Peak Name', 'purge', 'Rel.Area_x', 'Rel.Area_y']]
    purge = purge.rename(columns={'Rel.Area_y': sample2 + ' (% a/a)', 'Rel.Area_x': 'isolated (% a/a)' })
    
   #Inversion: Enrichment > 1
    if (inv == True):
        purge['purge'] = 1 / purge['purge'] 

    return purge


def getPurgeRandom(sample1, sample2, data, inv=True, fraction = 0.8):

#     sel =  [sample1, sample2]

    #   get purity data in accessible (non-PIVOT)-format & reset index
    if isinstance(data.index, pd.core.indexes.multi.MultiIndex):
        data_Purity = data[['Rel.Area']]
    
    data_Purity = data
    data_Purity = data_Purity.reset_index()
    
#     Extract sample data
    sample1_data = data_Purity.loc[data_Purity['Sample Type'] == sample1]
    sample2_data = data_Purity.loc[data_Purity['Sample Type'] == sample2]

    sample1_data = sample1_data.sample(frac=fraction, replace=False)
    sample2_data = sample2_data.sample(frac=fraction, replace=False)


# merge sample data: Match Experiments and Peak Names
    purge = pd.merge(sample1_data, sample2_data, on=['Experiment', 'Peak Name'], how='inner')
    purge

    purge['purge'] = purge['Rel.Area_x'] / purge['Rel.Area_y']

    purge = purge[['Experiment', 'Peak Name', 'purge']]
    
   #Inversion: Enrichment > 1
    if (inv == True):
        purge['purge'] = 1 / purge['purge'] 

    
    purgeAverage = purge.groupby('Peak Name').mean()
    
    
    return purgeAverage






def getPurgeMatrix(sam1, data, inv=True):
#Extract all purge factors relative to the selected sample sam1 in Sample_Type of data


    #   get  PIVOT-data into accessible (non-PIVOT)-format & reset index
    if isinstance(data.index, pd.core.indexes.multi.MultiIndex):
        data_Purity = data[['Rel.Area']]
    
    data_Purity = data
    data_Purity = data_Purity.reset_index()
    
    #Remove rare samples from list
    data_Purity = data_Purity.loc[(~data_Purity['Sample Type'].str.endswith(')'))&
                                  (~data_Purity['Sample Type'].str.contains('nnt'))&
                                  (~data_Purity['Sample Type'].str.contains('Ref'))
                                  ,:]    


    samList = data_Purity['Sample Type'].drop_duplicates()
    impList = data_Purity['Peak Name'].drop_duplicates()


    pMname= 'Purge_'+ sam1
    purgeM = pd.DataFrame(index=impList, columns=samList).astype(float)
    purgeM.columns.name=pMname


    current = pd.DataFrame()
    current['Peak Name'] = impList

    #     Extract sample data
    for i in samList: 
        add = getPurge(sam1, i, data_Purity, inv) 
        current = pd.merge(current, add, how='outer')
        current[i] = current['purge']
        current = current.drop(columns='purge')
        
        
#     average = current.groupby('Peak Name').mean()
#     worst = current.groupby('Peak Name').max()
#     best = current.groupby('Peak Name').min()
    
    
        
    return current

def getIPCspec(purgeData, IPCsample, prodSPEC, estimator = 'median', unknown_limit = ''):
    
    sample = IPCsample

    #  default = median
    purgeMatrix = purgeData.median().dropna(how='all')

    if (estimator == 'mean'):
        purgeMatrix = purgeData.mean().dropna(how='all')

    if (estimator == 'med'):
        purgeMatrix = purgeData.median().dropna(how='all')   

    if (estimator == 'median'):
        purgeMatrix = purgeData.median().dropna(how='all') 

    if (estimator == 'upper'):
        purgeMatrix = purgeData.quantile([0.75]).dropna(how='all')

    if (estimator == 'lower'):
        purgeMatrix = purgeData.quantile([0.25]).dropna(how='all')

        
    purgeM = purgeMatrix

    prodSPEC['USL'] = pd.to_numeric(prodSPEC['USL'], errors='coerce')
    prodSPEC['LSL'] = pd.to_numeric(prodSPEC['LSL'], errors='coerce')

    specUSL = prodSPEC[prodSPEC['USL'] < 100]
    specUSL = specUSL.set_index('Peak Name')

    specLSL = prodSPEC[prodSPEC['LSL'] > 0] 
    specLSL = specLSL.set_index('Peak Name')

    smUSL = pd.merge(specUSL[['USL']], purgeMatrix, on=['Peak Name'], how='inner')
    smLSL = pd.merge(specLSL[['LSL']], purgeMatrix, on=['Peak Name'], how='inner')


    smUSL = smUSL.apply(lambda x: x*x['USL'], axis=1).drop('USL', axis=1)
    smLSL = smLSL.apply(lambda x: x*x['LSL'], axis=1).drop('LSL', axis=1)

    #Fill impurites without purge factor by max tolerated level in isolated product
    noPF = float(100 - prodSPEC[prodSPEC['LSL'] > 0]['LSL'])
    unkUSL = 0.15  #Conservative!

    smUSL = smUSL.fillna(unkUSL)
     

    if unknown_limit == 'product':
            smUSL = smUSL.fillna(noPF)
            print('Limit without PF: ', noPF)
    else:
            print('Limit without PF: ', unkUSL)

    ipcSPEC = prodSPEC.drop(columns=['USL', 'LSL'])

    ipcSPEC = pd.merge(ipcSPEC, smLSL[[sample]], on='Peak Name', how = 'left')
    ipcSPEC = ipcSPEC.rename({sample:'LSL'},axis=1)
    ipcSPEC = pd.merge(ipcSPEC, smUSL[[sample]], on='Peak Name', how = 'left')
    ipcSPEC = ipcSPEC.rename({sample:'USL'},axis=1)


    ipcSPEC['LSL'] = ipcSPEC['LSL'].fillna(0)
    ipcSPEC.loc[ipcSPEC['LSL'] > 0, 'USL'] = 100

    ipcSPEC = ipcSPEC.dropna()

    return ipcSPEC, purgeM