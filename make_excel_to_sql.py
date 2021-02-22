import sqlite3
import pandas as pd 
from sqlalchemy import create_engine

file = 'one.xlsx'

engine = create_engine('sqlite://', echo=False)
df = pd.read_excel(file,sheet_name='Sheet1')

df.to_sql('employees',engine,if_exists='replace',index=False)


df5 = pd.read_sql_query("select COUNT(*) as v,depot from employees where MATERIAL='FTB15801' ",engine)

# final = pd.DataFrame(result)

print(df5)