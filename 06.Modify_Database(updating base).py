import pandas as pd
import math
df=pd.read_excel(r"C:\Users\sina.kian\Desktop\4\Results.xlsx")
dm=pd.read_excel(r"C:\Users\sina.kian\Desktop\4\ds.xlsx")
dff=df.copy()
for i in dm.index:
    a_row = pd.Series(dm.loc[i],index=df.columns)
    row_df = pd.DataFrame([a_row])
    df = pd.concat([df,row_df])

OH=[]
c_i=0
for i in dm['Locations']:
    c_j=0
    for j in dff['Locations']:
        if (i==j) and (dm['Company'][c_i]==dff['Company'][c_j]):
            df['Validity'][c_j]='Inv'+str(len(dff)+c_i-1)
            df['Validity'][len(dff)+c_i-1]='V'+str(c_j)
        c_j+=1
    c_i+=1
    OH=[]

OH=[]
df.index=range(0,len(df))
df=df[['Validity', 'update', 'date', 'N', 'Company', 'comp.(code)', 'Locations', 'site', 'state','country', 'zip code', 'Longitude', 'Latitude', 'accuracy', 'site type', 'Active/Deactive', 'Activity', 'Rental / Property', 'third party use', 'Area', 'building', 'machinery', 'building & machinery', 'stock', 'improvement of third party goods', 'molds', 'stock/content', 'Total PD', 'contr. margin', 'Total', 'Currency']]
file_name='dfff.xlsx'
df.to_excel(file_name)
