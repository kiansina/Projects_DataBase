#delet enter in columns' heads referred to as "Country (risk location)" , "Buildings (reconstruction value)", "Contents (replacement value)", "Stock (peak value per year)"
# Currencies and dates must be added manually
import pandas as pd
import math

# insert path/date/Currency/output/change excel columns
#df=pd.read_excel(r"C:\Users\sina.kian\Desktop\Work\Ariston Marco\Results.xlsx")
dp=pd.read_excel(r"C:\Users\sina.kian\Desktop\6\Ariston_Property_Location_2022 - Completed.xlsx")
#df1=df

dp.columns=dp.loc[10]
L=len(dp['Company name'].dropna())
dp=dp[11:11+L-1]
dp.index=range(0,L-1)
dm=pd.DataFrame()
dm['Validity']=["V"]*len(dp)
dm['update']=[1]*len(dp)
dm['date']=['08-02-2022']*len(dp)
dm['N']=[0]*len(dp)
dm['site type']=['-']*len(dp)
dm['Active/Deactive']=['-']*len(dp)
dm['Rental / Property']=['-']*len(dp)
dm['third party use']=['-']*len(dp)
dm['building & machinery']=[0]*len(dp)
dm['contr. margin']=[0]*len(dp)
dm['improvement of third party goods']=[0]*len(dp)
dm['molds']=[0]*len(dp)
dm['stock/content']=[0]*len(dp)
dm['Currency']=['GBP']*len(dp)

dm['Company']=['-']*len(dp)
dm['comp.(code)']=['-']*len(dp)
dm['Locations']=['-']*len(dp)
dm['site']=['-']*len(dp)
dm['state']=['-']*len(dp)
dm['country']=['-']*len(dp)
dm['zip code']=['-']*len(dp)
dm['Longitude']=['-']*len(dp)
dm['Latitude']=['-']*len(dp)
dm['Activity']=['-']*len(dp)
dm['Area']=['-']*len(dp)
dm['building']=['-']*len(dp)
dm['machinery']=['-']*len(dp)
dm['stock']=['-']*len(dp)
dm['accuracy']=['-']*len(dp)

dm['Total']=['-']*len(dp)
dm['Total PD']=['-']*len(dp)

for i in range(0,L-1):
    dm['N'][i]=len(df1)+i
    if dp['Active (yes/not)'][i]=='Yes':
        dm['Active/Deactive'][i]=1
    elif dp['Active (yes/not)'][i]=='No':
        dm['Active/Deactive'][i]=2
    else:
        dm['Active/Deactive'][i]=0
    if dp['Is your Company the Owner of the location ? (YES/NO)'][i]=='YES':
        dm['Rental / Property'][i]='property'
    elif dp['Is your Company the Owner of the location ? (YES/NO)'][i]=='NO':
        dm['Rental / Property'][i]='rental'
    else:
        dm['Rental / Property'][i]='0'
    if dp['Is your Company the Operator of the location ? (YES/NO)'][i]=='YES':
        dm['third party use'][i]='no'
    elif dp['Is your Company the Operator of the location ? (YES/NO)'][i]=='NO':
        dm['third party use'][i]='yes'
    else:
        dm['third party use'][i]='0'

dp['Latitude']=dp['Latitude'].fillna('-')
dp['Longitude']=dp['Longitude'].fillna('-')
for i in range(0,L-1):
    dm['Company'][i]=dp['Company name'][i]
    dm['comp.(code)'][i]=dp['Company code'][i]
    dm['Locations'][i]=dp['Address'][i]
    dm['site'][i]=dp['City'][i]
    dm['state'][i]=dp['State or Region'][i]
    dm['country'][i]=dp['Country(risk location)'][i]
    dm['zip code'][i]=dp['ZIP Code'][i]
    dm['Longitude'][i]=dp['Longitude'][i]
    dm['Latitude'][i]=dp['Latitude'][i]
    dm['Activity'][i]=dp['Main occupation'][i]
    dm['Area'][i]=dp['Covered Area (sqm)'][i]
    dm['building'][i]=dp['Buildings(reconstruction value)'][i]
    dm['machinery'][i]=dp['Contents(replacement value)'][i]
    dm['stock'][i]=dp['Stock(peak value per year)'][i]
    if dm['Latitude'][i]=='-':
        dm['accuracy'][i]='-'
    else:
        dm['accuracy'][i]=1

dm['Total PD'] = dm.loc[0 : L-1,['building', 'machinery','stock']].sum(axis = 1)
dm=dm[['update', 'date', 'N', 'Company', 'comp.(code)', 'Locations', 'site', 'state','country', 'zip code', 'Longitude', 'Latitude', 'accuracy', 'site type', 'Active/Deactive', 'Activity', 'Rental / Property', 'third party use', 'Area', 'building', 'machinery', 'building & machinery', 'stock', 'improvement of third party goods', 'molds', 'stock/content', 'Total PD', 'contr. margin', 'Total', 'Currency']]

dn=pd.DataFrame()
#ds_row = pd.Series([0]*30,index=dm.columns)
ds_row = pd.Series(index=dm.columns)
ds = pd.DataFrame([ds_row])
ds.columns=dm.columns
oh=[]
for i in range(0,len(dm)):
    if dm['Locations'][i] not in oh:
        dl=dm[dm['Locations']==dm['Locations'][i]]
        oh.append(dm['Locations'][i])
        ohh=[]
        for jj in dl.index:
            if dl['Company'][jj] not in ohh:
                dll=dl[dl['Company']==dl['Company'][jj]]
                ohh.append(dl['Company'][jj])
                a_row = pd.Series(dll.loc[jj],index=dm.columns)
                row_df = pd.DataFrame([a_row])
                ds = pd.concat([ds,row_df])
                A=dll['Activity'].unique()
                AAA=[]
                for mm in A:
                    if mm not in AAA:
                        AAA.append(mm)
                print(dll[['machinery','stock']])
                ds['building'][jj]=dll['building'].sum()
                ds['machinery'][jj]=dll['machinery'].sum()
                ds['stock'][jj]=dll['stock'].sum()
                ds['Total PD'][jj]=dll['Total PD'].sum()
                ds['Area'][jj]=dll['Area'].sum()
                ds['Activity'][jj]=AAA
                print(ds[['machinery','stock']])


ds.index=range(0,len(ds))
ds=ds[1:]
ds.index=range(0,len(ds))
ds['Validity']=['V']*len(ds)
ds=ds[['Validity', 'update', 'date', 'N', 'Company', 'comp.(code)', 'Locations', 'site', 'state','country', 'zip code', 'Longitude', 'Latitude', 'accuracy', 'site type', 'Active/Deactive', 'Activity', 'Rental / Property', 'third party use', 'Area', 'building', 'machinery', 'building & machinery', 'stock', 'improvement of third party goods', 'molds', 'stock/content', 'Total PD', 'contr. margin', 'Total', 'Currency']]







file_name='ds.xlsx'
ds.to_excel(file_name)
