import pandas as pd
import numpy as np
import xlsxwriter
import re


########### Create a Pandas dataframe from the data.
############ Create a Pandas Excel writer using XlsxWriter as the engine.
########### Convert the dataframe to an XlsxWriter Excel object.
source_data =('/Users/junkim/Desktop/Workbook1.csv')
writer = pd.ExcelWriter('/Users/junkim/Desktop/HELLO.xlsx', engine='xlsxwriter')



df = pd.read_csv(source_data, sep=',', header=None, low_memory=False)

df.columns = ["ID", "LOCATION"]

df1 = df

df1["LOCATION"] = df["LOCATION"]
df1['BUILD'] = df1['LOCATION'].str.extract('(-......|-....-.|-....|-...-..|-...-.|-...|-..-.|-..|-.-.|-.)', expand=True)





df1["ROOMnumber"] = df1["LOCATION"]




df1.columns = ["ID", "LOCATION", "VA_BUILD", "VA_ROOMnumber"]

df1




df1.to_excel(writer, sheet_name= 'LogData')
workbook = writer.book
#df3.to_excel(writer, sheet_name= 'bob')

##########Excel Sheets
workbook = writer.book
dataframeworksheet = writer.sheets['LogData']
dataframeworksheet.set_column('A:A',10)
dataframeworksheet.set_column('A:B',12)
dataframeworksheet.set_column('A:C',20)
bold = workbook.add_format({'bold': True})

dataframeworksheet.conditional_format('H2:H1000', {'type': 'text',
	'criteria': 'containing',
	'value': "Error",
	'format': bold})


writer.save()
