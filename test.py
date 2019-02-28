import os
import xlrd
import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import mysql.connector
from mysql.connector import errorcode

config = {
    'host':'finengdb.crjg5e9ji3h7.us-east-2.rds.amazonaws.com',
    'user':'am3p',
    'password': 'measure0',
    'database': 'finengdb'
}

try:
   conn = mysql.connector.connect(**config)
   print("Connection established")
except mysql.connector.Error as err:
  if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
    print("Something is wrong with the user name or password")
  elif err.errno == errorcode.ER_BAD_DB_ERROR:
    print("Database does not exist")
  else:
    print(err)
else:
  cursor = conn.cursor()

wrbk = load_workbook(filename='kospi_close.xlsx', read_only=True)
wrst = wrbk['종가']

df = pd.DataFrame(wrst.values)

symbol = df.iloc[8,:].tolist()
name = df.iloc[9,:].tolist()
data = df.iloc[14:,:]



# for i in range(1, len(symbol)):
#     nm = name[i]
#     symb = symbol[i]
#     for j in range(data.shape[0]):
#     # for j in range(10):
#         if data.iloc[j,i] is not None:
#             sql = "INSERT INTO kospi (dt, ticker, name, close) VALUES " \
#                   "('{}', '{}', '{}', {})".format(data.iloc[j,0]._short_repr, symb, nm, data.iloc[j,i])
#             print(sql)
#             cursor.execute(sql)
#     conn.commit()