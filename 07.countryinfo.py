import pandas as pd
from countryinfo import CountryInfo
df=pd.read_excel(r'"C:\Users\sina.kian\Desktop\work2\Dati_Paesi.xlsx"')
df.columns=df.loc(5)
df=df[6:]
dm=pd.DataFrame()
dm['Country']=df['Country/Territory']
dm['Capital']=[0]*len(dm)
dm['latitude']=[0]*len(dm)
dm['longitude']=[0]*len(dm)
df.index=range(0,len(dm))
for i in range(0,len(dm)+1):
    try:
        dm['Capital'][i]=CountryInfo(dm['Country'][i]).capital()
        dm['latitude'][i]=CountryInfo(dm['Country'][i]).capital_latlng()[0]
        dm['longitude'][i]=CountryInfo(dm['Country'][i]).capital_latlng()[1]
    except:
        dm['Capital'][i]=0
        dm['latitude'][i]=0
        dm['longitude'][i]=0
