#finding the address and giving the coordinates
import pandas as pd
import math
from functools import partial
from geopy.geocoders import Nominatim
def isnan(value):
    try:
        return math.isnan(float(value))
    except:
        return False

geolocator=Nominatim(user_agent="googleearth")
geocode=partial(geolocator.geocode,language="en")
df = pd.read_excel(r'C:\Users\sina.kian\Desktop\work2\Ubicazioni_Ariston.xlsx')
df.columns=df.iloc[3]
df=df.iloc[4:,:]
Ad=[]
for i in range(4,277):
    if not isnan(df['nazione'][i]):
        A=str(df['nazione'][i])

    if not isnan(df['città'][i]):
        A=A+', '+str(df['città'][i])

    if not isnan(df['indirizzo'][i]):
        A=A+', '+str(df['indirizzo'][i])

    if not isnan(df['n°'][i]):
        A=A+' '+str(df['n°'][i])

    Ad.append(A)

Longi=[]
Lati=[]
le=[]
tr=[]
C=1
for i in Ad:
    print(C)
    C+=1
    location=geolocator.geocode(i)  #location=geolocator.geocode(i,timeout=10)
    le.append(len(i))
    t=0
    while location is None:
        if len(i)>5:
            i=i[0:len(i)-1]
            t+=1
            location=geolocator.geocode(i)
        else:
            Longi.append(0)
            Lati.append(0)
    else:
        tr.append(t)
        Longi.append(location.longitude)
        Lati.append(location.latitude)

dff=pd.DataFrame()
dff['Adress']=Ad
dff["Longi"]=Longi
dff["Lati"]=Lati
dff["precision"]=[(x1 - x2)/x1 for (x1, x2) in zip(le, tr)]
dff["Length"]=le
dff["number of tries"]=tr

file_name='victoryha.xlsx'
dff.to_excel(file_name)
