#Before starting replace '\ ' with '\' in 'Items' column.
import pandas as pd
import re
import math
df=pd.read_excel(r"C:\Users\sina.kian\Desktop\code debug\New folder\Property Proposta allocazione **** 20.01.22_UK in FOS VER8.xlsx")

#organize columns and proper length

df.columns=df.iloc[0]
df=df[1:]
df=df.dropna(subset='Country')
#df=df.iloc[:369]
LENG=len(df)
pp=[]
#Allocate a index number as a column
for i in range(1,LENG+1):
    j=str(i)
    pp.append(j)

df['number']=pp #or df['number']=df.index

df['Locations'] = df['Locations'].fillna(df['number'])

#Cleaning Items column:
def isnan(value):
    try:
        import math
        return math.isnan(float(value))
    except:
        return False


items=['-']*len(df['Items'])
for i in range(1,len(df['Items'])+1):
    if type(df['Items'][i])==str:
        items[i-1]=df['Items'][i].strip().lower()

df['Items']=items
df[df['Items']!='-']['Items'].unique()
#below line should be deleted
#df=df[:148]

dff=df[['Locations','Activity', 'Items', 'Rental / Property','PD&BI Values 2021','Country', 'Comp. (code)', 'Company (name)', 'Site', 'Site type', 'number']]
dff=dff.sort_values(by=['Locations'])
dff.index=range(0,len(dff))

#Creation of favorite dataframe

dm=pd.DataFrame()
dm['Locations']=dff['Locations'].unique()
dm['N']=dm.index
dm['Activity']=[0]*len(dm)
dm['Company']=[0]*len(dm)

OH=[]
c_i=0
for i in dm['Locations']:
    c_j=0
    for j in dff['Locations']:
        if i==j:
            if dm['Company'][c_i]==0:
                dm['Company'][c_i]=dff['Company (name)'][c_j]
            else:
                if dm['Company'][c_i]==dff['Company (name)'][c_j]:
                    pass
                elif (dm['Company'][c_i]!=dff['Company (name)'][c_j]) and (dff['Company (name)'][c_j] not in OH):
                    a_row = pd.Series([i, -10, dm['Activity'][c_i], dff['Company (name)'][c_j]],index=dm.columns)
                    OH.append(dff['Company (name)'][c_j])
                    row_df = pd.DataFrame([a_row])
                    dm = pd.concat([row_df, dm])
        c_j+=1
    c_i+=1
    OH=[]

OH=[]




dm=dm.sort_values(by=['Locations'])
dm.index=range(0,len(dm))




OH=[]
c_i=0
for i in dm['Locations']:
    c_j=0
    for j in dff['Locations']:
        if (i==j) and (dm['Company'][c_i]==dff['Company (name)'][c_j]):
            if dm['Activity'][c_i]==0:
                dm['Activity'][c_i]=dff['Activity'][c_j]
            else:
                if (dm['Activity'][c_i]==dff['Activity'][c_j]) and (dm['Company'][c_i]==dff['Company (name)'][c_j]):
                    pass
                elif (dm['Activity'][c_i]!=dff['Activity'][c_j]) and (dm['Company'][c_i]==dff['Company (name)'][c_j]):
                    try:
                        math.isnan(float(dff['Rental / Property'][c_j]))
                        pass
                    except:
                        if ((isnan(dm['Activity'][c_i])) or (isnan(dm['Activity'][c_j])))==False:
                            dm['Activity'][c_i]=dm['Activity'][c_i]+'/'+dff['Activity'][c_j]
        c_j+=1
    c_i+=1

dm['site type']=['-']*len(dm)
dm['site']=['-']*len(dm)
dm['Rental / Property']=['-']*len(dm)
c_i=0
for i in dm['Locations']:
    c_j=0
    for j in dff['Locations']:
        if (i==j) and (dm['Company'][c_i]==dff['Company (name)'][c_j]):
            if dm['site type'][c_i]=='-':
                dm['site type'][c_i]=dff['Site type'][c_j]
            else:
                if dm['site type'][c_i]==dff['Site type'][c_j]:
                    pass
                else:
                     try:
                         math.isnan(float(dm['site type'][c_i]))
                         pass
                     except:
                         dff['Site type'][c_j] not in dm['site type'][c_i].split('/')
                         try:
                             math.isnan(float(dff['Rental / Property'][c_j]))
                             pass
                         except:
                             dm['site type'][c_i]=dm['site type'][c_i]+'/'+dff['Site type'][c_j]
            if dm['site'][c_i]=='-':
                dm['site'][c_i]=dff['Site'][c_j]
            else:
                if dm['site'][c_i]==dff['Site'][c_j]:
                    pass
                elif dff['Site'][c_j] not in dm['site'][c_i].split('/'):
                    try:
                        math.isnan(float(dff['Rental / Property'][c_j]))
                        pass
                    except:
                        dm['site'][c_i]=dm['site'][c_i]+'/'+dff['Site'][c_j]
            if dm['Rental / Property'][c_i]=='-':
                try:
                    math.isnan(float(dff['Rental / Property'][c_j]))
                    pass
                except:
                    dm['Rental / Property'][c_i]=dff['Rental / Property'][c_j]
        c_j+=1
    c_i+=1

dm['country']=['-']*len(dm)
c_i=0
for i in dm['Locations']:
    c_j=0
    for j in dff['Locations']:
        if (i==j) and (dm['Company'][c_i]==dff['Company (name)'][c_j]):
            if dm['country'][c_i]=='-':
                dm['country'][c_i]=dff['Country'][c_j]
        c_j+=1
    c_i+=1

dm=dm.sort_values(by=['country','Locations'])
dm.index=range(0,len(dm))
dm['comp.(code)']=['-']*len(dm)
c_i=0
for i in dm['Company']:
    c_j=0
    for j in dff['Company (name)']:
        if i==j:
            if dm['comp.(code)'][c_i]=='-':
                dm['comp.(code)'][c_i]=dff['Comp. (code)'][c_j]
        c_j+=1
    c_i+=1


dm['building']=['-']*len(dm)
dm['machinery']=['-']*len(dm)
dm['building & machinery']=['-']*len(dm)
dm['stock']=['-']*len(dm)
dm['contr. margin']=['-']*len(dm)
dm['improvement of third party goods']=['-']*len(dm)
dm['molds']=['-']*len(dm)
dm['stock/content']=['-']*len(dm)

dn = dm.set_index(['Locations', 'Company',dm.index])
SS=0
for A in dn.index:
    dh=dff[(dff['Locations']==A[0]) & (dff['Company (name)']==A[1])][['Locations','Company (name)','Items','PD&BI Values 2021']]
    MN=0
    MN1=0
    MN2=0
    MN3=0
    MN4=0
    MN5=0
    MN6=0
    MN7=0
    for ZZ,MM in dh[['Items','PD&BI Values 2021']].itertuples(index=False):
        if ZZ=='buildings' or ZZ=='pref. building' or ZZ=='telonati' or ZZ=='office lease' or ZZ=='office/warehouse lease':
            MN+=MM
        if ZZ=='machinery' or ZZ=='machinery/equipment' or ZZ=='equipment' or ZZ=='contents':
            MN1+=MM
        if ZZ=='stock':
            MN2+=MM
        if ZZ=='building & machinery' or ZZ=='equipment, workplace':
            MN3+=MM
        if ZZ=='contr. margin':
            MN4+=MM
        if ZZ=='improvement of third party goods':
            MN5+=MM
        if ZZ=='molds':
            MN6+=MM
        if ZZ=='stock/content':
            MN7+=MM
    dm['building'][A[2]]=MN
    dm['machinery'][A[2]]=MN1
    dm['stock'][A[2]]=MN2
    dm['building & machinery'][A[2]]=MN3
    dm['contr. margin'][A[2]]=MN4
    dm['improvement of third party goods'][A[2]]=MN5
    dm['molds'][A[2]]=MN6
    dm['stock/content'][A[2]]=MN7

dm=dm.sort_values(by=['Locations'])
dm.index=range(0,len(dm))
dh=dm[['N', 'Company', 'comp.(code)', 'Locations', 'site', 'country', 'site type', 'Activity', 'Rental / Property', 'building', 'machinery', 'building & machinery', 'stock', 'contr. margin', 'improvement of third party goods', 'molds', 'stock/content']]
dh['Total PD']= dh.loc[0 : len(dh),['building', 'machinery', 'building & machinery', 'stock', 'improvement of third party goods', 'molds', 'stock/content']].sum(axis = 1)
dh['Total'] = dh.loc[0 : len(dh),['building', 'machinery', 'building & machinery', 'stock', 'contr. margin', 'improvement of third party goods', 'molds', 'stock/content']].sum(axis = 1)
dh=dh[['N', 'Company', 'comp.(code)', 'Locations', 'site', 'country', 'site type', 'Activity', 'Rental / Property', 'building', 'machinery', 'building & machinery', 'stock', 'improvement of third party goods', 'molds', 'stock/content', 'Total PD', 'contr. margin', 'Total']]
#Add columns:
a_row = pd.Series(['-','-','-','-','-','-','-','-','-', dh['building'].sum(), dh['machinery'].sum(), dh['building & machinery'].sum(), dh['stock'].sum(),  dh['improvement of third party goods'].sum(), dh['molds'].sum(), dh['stock/content'].sum(), dh['Total PD'].sum(), dh['contr. margin'].sum(), dh['Total'].sum()], index=dh.columns)
row_df = pd.DataFrame([a_row])
dh = pd.concat([dh,row_df])

#data['total'] = data['age'].sum()


df=dh

df1=df[0:len(df)-1]
df1['update']=[0]*len(df1)
df1['date']=['31-12-21']*len(df1)
df1['Active/Deactive']=[0]*len(df1)
df1['state']=['-']*len(df1)
df1['zip code']=['-']*len(df1)
df1['third party use']=['-']*len(df1)
df1['Area']=['-']*len(df1)
df1['Longitude']=['-']*len(df1)
df1['Latitude']=['-']*len(df1)
df1['Currency']=['Euro']*len(df1)
df1['accuracy']=['-']*len(df1)
df1['Validity']=["V"]*len(df1)
df1=df1[['Validity','update', 'date', 'N', 'Company', 'comp.(code)', 'Locations', 'site', 'state','country', 'zip code', 'Longitude', 'Latitude', 'accuracy', 'site type', 'Active/Deactive', 'Activity', 'Rental / Property', 'third party use', 'Area', 'building', 'machinery', 'building & machinery', 'stock', 'improvement of third party goods', 'molds', 'stock/content', 'Total PD', 'contr. margin', 'Total', 'Currency']]
for i in df1.index:
    df1['Company'][i]=df1['Company'][i].capitalize()
    df1['Company'][i]=re.sub('\.', '',df1['Company'][i])
    df1['Company'][i]=re.sub(',', '',df1['Company'][i])
    df1['Locations'][i]=df1['Locations'][i].capitalize()
    df1['Locations'][i]=re.sub('\.', '',df1['Locations'][i])



file_name='Resultstest.xlsx'

df1.to_excel(file_name)
