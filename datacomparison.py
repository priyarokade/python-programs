
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import xlsxwriter

df1 = pd.read_excel('E:\Priya\Python\Practice\Citi_Industrialization\CP_AGGR.xlsx', sheet_name='CP_AGGR')
df2 = pd.read_excel('E:\Priya\Python\Practice\Citi_Industrialization\CP_REVENUE_PACK.xlsx', sheet_name='CP_REVENUE_PACK')
print(df1)
print(df2)
df1['key1'] =""
df2['key2'] =""

key_input = input("Enter a key element separated by space ")
keylist  = key_input.split()

results_col_input = input("Enter a result column list  : ")
resultlist  = results_col_input.split()


for x in keylist:
    print(x)


for i , row in df1.iterrows():
        keyvalue = ""
        for x in keylist :
            keyvalue = keyvalue + str(row[x])
            print(keyvalue.replace(" ",""))
        df1.at[i, 'key1'] = keyvalue.replace(" ","")

print(df1)

for i , row in df2.iterrows():
        keyvalue = ""
        for x in keylist :
            keyvalue = keyvalue + str(row[x])
            print(keyvalue.replace(" ",""))
        df2.at[i, 'key2'] = keyvalue.replace(" ","")

print(df2)

resultset = pd.merge(df1,df2,left_on='key1',right_on='key2')
"""
for x in resultlist:
    newresultcolname = x + '_diff'
    resultset[newresultcolname] = ""
"""

print(resultset)

for i , row in resultset.iterrows():
    for x in resultlist:
        newresultcolname = x + '_diff'
        check_x_colname = newresultcolname.replace("_diff", "_x")
        check_y_colname = newresultcolname.replace("_diff", "_y")
        diff_value = row[check_x_colname] - row[check_y_colname]
        print("value at %d row and colume %s : %d " %(i,newresultcolname, diff_value) )
        resultset.at[i, newresultcolname] = diff_value

print(resultset)


writer = pd.ExcelWriter('E:\Priya\Python\Practice\Citi_Industrialization\Result.xlsx', engine='xlsxwriter')
resultset.to_excel(writer, sheet_name="Sheet1", index=False)
writer.save()