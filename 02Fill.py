import pandas as pd
import numpy as np
import re

import 01Empty_DB as EDB
Group=EDB.Group
Company=EDB.Company
Locations=EDB.Locations
Transportation=EDB.Transportation
Prodotto=EDB.Prodotto
Revenue_Country=EDB.Revenue_Country
Dependenza_Clienti=EDB.Dependenza_Clienti
Sinistri=EDB.Sinistri


df = pd.read_excel(r"C:\Users\sina.kian\Desktop\Quick report format V2\SRC_Questionario Base_Integrale2020_****.xlsx",sheet_name=None)

Gruppo=df['COP'].loc[6][0]
df1=df['1.DB Società']
df2=df['Criteri']
df3=df['2.DB ubicazioni']
df4=df['3.Fatturati Totali']
df5=df['3.Fatturati Settore']
df6=df['3.Attività']
df7=df['3.Processi']
df8=df['4-5.Antincendio Antintrusione']
df9=df['Registro_Rischi']
df10=df['6.Merci']
df11=df['8.Trasporto IN']
df12=df['8.Trasporto OUT']
df13=df['7.Sinistri']
df14=df['9.HR']
df15=df['Flussi']
df16=df['MCr1']
df17=df['MCr2']
df18=df['MCr3']
df19=df['MCr4']
df20=df['Flussi_Mappa']
df21=df['Servizio']
#°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°      df1 ---> Company
cols1=df1.loc[1]
cols1_1=cols1[1:4]
cols1_3=cols1[10:15]
cols=df1.loc[2]
cols1_2=cols[4:10]
COLS=[]
COLS[0:3]=cols1_1
COLS[3:9]=cols1_2
COLS[10:15]=cols1_3
COLS=['n']+COLS
df1.columns=COLS
df1=df1[3:]

if Gruppo not in Group['Group Name']:
    if len(Group['Group Name'])==0:
        GG= pd.DataFrame([[1,Gruppo]], columns=Group.columns)
        Group=pd.concat([Group, GG])
    else:
        GG= pd.DataFrame([[Group['Group ID'].max()+1,Gruppo]], columns=Group.columns)
        Group=pd.concat([Group, GG])
else:
    print("The Group is already in database")

df1['Group ID']=list(Group[Group['Group Name']==Gruppo]['Group ID'])*len(df1)
Start=len(Company)
df1['Company ID']=range(Start+1,Start+len(df1)+1)
df1=df1[['Company ID','Denominazione societaria ','Group ID', '% Part.', 'Tramite', 'Indirizzo','n.c.', 'cap', 'città', 'pv', 'nazione', 'Area di business','Principale attività svolta', 'Capitale sociale', 'Partita IVA','Note']]
df1.columns=['Company ID','Company Name','Group ID','% Part.', 'Tramite', 'Indrizzo SS', 'n.c. SS', 'cap SS', 'Site SS', 'pv SS','nazione SS','Area of Business','PAS','CS','P.IVA', 'Note']
A=Company.merge(df1,how='outer')
Company=A[Company.columns]
############################################################### Company First fill finished
#°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°      df2 ---> Criteri_probabilità & Criteri_Impatto
df2_1=df2.copy()
df2_1.columns=df2_1.loc[1]
df2_1=df2_1[2:7]
df2_1.index=range(0,len(df2_1))
df2_2=df2.copy()
df2_2.columns=df2_2.loc[9]
df2_2=df2_2[10:15]
df2_2.index=range(0,len(df2_2))
df2_1=df2_1[['Livello', 'Descrizione']]
df2_1.columns=['Livello Quantitativo','Livello Qualitativo','Descrizione']
Criteri_probabilità=df2_1
df2_2=df2_2[['Livello','Quantitativo','Qualitativo']]
df2_2.columns=['Livello Quantitativo','Livello Qualitativo','Descrizione Quantitativo','Descrizione Qualitativo']
Criteri_Impatto=df2_2

################################################################ Criteri_probabilità & Criteri_Impatto finished
#°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°      df3 ---> Locations
df3.columns=df3.iloc[3]
df3=df3[4:]
df3=df3[['società','indirizzo','n°','frazione/località/ specifiche','cap','città','pv','nazione','link mappa/ latitudine e longitudine','Anno costruzione','bene culturale (SI/NO)','soggetto proprietario','superficie di sviluppo totale (mq) (1)','superficie di sviluppo specifica (mq) (1)','destinazione d\'uso generale (2)','destinazione d\'uso specifica (3)','N. Dipendenti','Ricavi infragruppo','Autorizzazioni ambientali AIA o AUA','Fabbricato','Contenuto (5)','Merci\n(max giacenza)','Totale','Polizza']]
df3.dropna(subset=['società'],inplace=True)
df3.index=range(0,len(df3))
df3['Complete Address ID']=range(1,len(df3)+1)
df3=df3[['Complete Address ID','società','indirizzo','n°','frazione/località/ specifiche','cap','città','pv','nazione','link mappa/ latitudine e longitudine','Anno costruzione','bene culturale (SI/NO)','soggetto proprietario','superficie di sviluppo totale (mq) (1)','superficie di sviluppo specifica (mq) (1)','destinazione d\'uso generale (2)','destinazione d\'uso specifica (3)','N. Dipendenti','Ricavi infragruppo','Autorizzazioni ambientali AIA o AUA','Fabbricato','Contenuto (5)','Merci\n(max giacenza)','Totale','Polizza']]
cols3=['Complete Address ID','Company','Address','n°','frazione località  specifiche','cap','City','pv','nazione','link mappa/ latitudine e longitudine','Construction Year','bene culturale','soggetto proprietario','SST[m2]','SSS[m2]','DUG','DUS','N. Dipendenti','Ricavi infragruppo','Pollution Classification','Fabbricato','Contenuto','Merci','Total Value','Polizza']
df3.columns=cols3
df3=df3[['Complete Address ID','Company','Address','n°','frazione località  specifiche','cap','City','pv','nazione','Construction Year','bene culturale','soggetto proprietario','SST[m2]','SSS[m2]','DUG','DUS','N. Dipendenti','Ricavi infragruppo','Pollution Classification','Fabbricato','Contenuto','Merci','Total Value','Polizza']]
##%%
C_Locations=['','','','','','','','','','','','','Longitude','Latitude','Accuracy','','Building material','','','','','','Site type','Activity situation','Activity Type','Building Ownership','Third party use','Number of Floors','','','','','','','','','Business interruption','Total including BI','','RM Max' ,'SF Max' ,'FP Max' ,'FPS Max' ,'TPP Max' ,'TOTAL products Max','Squadra interna emergenza','Permesso fumo','Piano emergenza','Esercitazione annuale','Pianificazione programmi manutenzione' ,'Spinkler','Idranti esterni','Manichette interne','Riserve acqua antincendio','esclusiva antincendio','elettriche','diesel','generatore emergenza','Rilevatori incendio/fumo','R.I.F Coefficient','R.I.F Category','Estintori','Sistemi speciali protezione','Cat SSP','Rilevatori perimetrali','Rilevatori microonde interni','Sistemi accesso controllato','Illuminazione su allarme','Sirena','Sirena Category','TVCC','Number of TVCC','TVCC Colori','TVCC Night Vision','registrazione immagini','Servizio vigilanza','V.F più vicina','distanza (km)','tempo intervento [min] ','dimensione riserva [m3] ','Destinatario allarme']

######
df3['Complete Address']=[-500]*len(df3)
df3['Company ID']=[-500]*len(df3)
for i in df3.index:
    df3['Complete Address'][i]=str(df3['nazione'][i])+' '+str(df3['City'][i])+' '+str(df3['Address'][i])+' '+str(df3['n°'][i])

for i in df3.index:
    df3['Complete Address'][i]=re.sub(' +', ' ',df3['Complete Address'][i])

df3['Group ID']=list(Group[Group['Group Name']==Gruppo]['Group ID'])*len(df3)

for i in df1.index:
    for j in df3.index:
        if df3['Company'][j]==df1['Company Name'][i]:
            df3['Company ID'][j]=df1['Company ID'][i]

B=Locations.merge(df3,how='outer')
Locations=B[Locations.columns]
################################################################ Locations finished

writer = pd.ExcelWriter('DB_StructureV2.xlsx', engine='xlsxwriter')
Group.to_excel(writer, sheet_name='Group')
Company.to_excel(writer, sheet_name='Company')
Locations.to_excel(writer, sheet_name='Locations')
Transportation.to_excel(writer, sheet_name='Transportation')
Prodotto.to_excel(writer, sheet_name='Prodotto')
Revenue_Country.to_excel(writer, sheet_name='Revenue_Country')
Dependenza_Clienti.to_excel(writer, sheet_name='Dependenza_Clienti')
Sinistri.to_excel(writer, sheet_name='Sinistri')
writer.save()
