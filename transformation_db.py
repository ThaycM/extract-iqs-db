import pandas as pd

# IMPORTS #
db_path=r"C:/Users/00071228/OneDrive - ENERCON/Pessoal/reports/IQS_Kaizen_MM/actions_db.xlsx"
df_db=pd.read_excel(db_path)
columns=['Status','Title','Description','Planned end','Actual end','Created on',
         'Linked with','Read','Editor','Creator','Activated at','% complete',
         'Action result','Actual begin', 'Completion status (bar)','Creator (position)',
         'Organization - editor','Partial task of','Planned begin','Version date']
df_db=df_db[[c for c in columns]]
print(df_db)
mascara_creator=~df_db["Creator"].isin(['Costa, Pedro Alexandre', 'Krause, Stefan'])
df_db=df_db[mascara_creator]
mascara_status=df_db["Status"].isin(['Completed','In progress','Unprocessed'])
df_db=df_db[mascara_status]
mascara_description=~df_db["Description"].isna()
df_db=df_db[mascara_description]
mascara_description_2=~df_db['Description'].str.lower().str.startswith('complaint process', na=False)
df_db=df_db[mascara_description_2]
print(df_db.info)