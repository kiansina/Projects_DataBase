import sys
import geopandas as gpd
import pandas as pd
import matplotlib.pyplot as plt



pd.concat([df1, df4.reindex(df1.index)], axis=1)

A={'C1':['A','B','C'],'C2':[1,2,3]}
AA=pd.DataFrame(A)
AA.set_index('C1',inplace=True)
B={'C4':['B','C','D'],'C5':[6,7,39]}
BB=pd.DataFrame(B)
BB.set_index('C4',inplace=True)
CC=pd.merge(AA,BB, left_index=True, right_index=True, how='left')
