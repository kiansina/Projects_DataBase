# streamlit_app.py
import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from gspread_pandas import Spread,Client
import ssl
ssl._create_default_https_context=ssl._create_unverified_context
# Create a connection object.
scope=["https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"]
credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scope,
)
client=Client(scope=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"], creds=credentials)
spreadsheetname="Questionnaire_test"
spread=Spread(spreadsheetname,client=client)

################
s=spread

s.sheets
s.url
s.open_sheet(0) #or the name of sheet

sh=client.open(spreadsheetname)
df=pd.DataFrame(sh.worksheet('test').get_all_records())
worksheet_list = sh.worksheets()
#df=s.sheet_to_df(header_rows=2).astype(int)
df=s.sheet_to_df(header_rows=1,index=False)

dx=pd.DataFrame(columns=['grade','number'])
dx['grade']=[20, 18,16]
dx['number']=[2, 5,3]

s.df_to_sheet(dx,sheet='another test', start='B1', replace=True, freeze_headers=1)
s.update_cells((1,1),(1,2),['created by', s.url])
