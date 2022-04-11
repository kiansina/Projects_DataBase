import requests
import pandas as pd
df = pd.read_excel(r'B:\Strategica organized\Data\RAPEX\Safety.Gate.weekly.reports_en.xlsx')
MM=df.values
for i in MM:
    if i[0]<2013:
        resp = requests.get(i[3])
        if i[1]>9:
            name=str(i[0])+str(i[1])+'.xlsx'
        else:
            name=str(i[0])+'0'+str(i[1])+'.xlsx'
        output = open(name, 'wb')
        output.write(resp.content)
        output.close()
