import pandas as pd
count=0
Year=[]
week=[]
df1 = pd.read_excel(r"C:\Users\sina.kian\Desktop\work\200806.xlsx")
df2 = pd.read_excel(r"C:\Users\sina.kian\Desktop\work\200746.xlsx")
col1=df1.iloc[1]
col2=df2.iloc[0]
for i in range (2005,2023):
    for j in range(1,54):
        if j>9:
            name=str(i)+str(j)+'.xlsx'
        else:
            name=str(i)+'0'+str(j)+'.xlsx'
        try:
            df = pd.read_excel(name)
            if df.iloc[1,1]=='Type':
                df.columns=df.iloc[1]
                df=df.iloc[2:,:]
            elif df.iloc[0,1]=='Type':
                df.columns=df.iloc[0]
                df=df.iloc[1:,:]

            if df.columns[5]!='Counterfeit':
                if len(df.columns)==26:
                    df.columns=col1
                elif len(df.columns)==24:
                    df.columns=col2
            if 'Company recall code (**)' not in df.columns:
                df.insert(23, 'Production dates (**)', ['N']*len(df))
                df.insert(23, 'Company recall code (**)', ['N']*len(df))
            Year=[i]*len(df)
            Week=[j]*len(df)
            df['Year']=Year
            df['Week']=Week
            count=count+1
            if count==1:
                dff=df
            else:
                dff=dff.append(df)
        except:
            continue

file_name='RAPEX.xlsx'
dff.to_excel(file_name)
