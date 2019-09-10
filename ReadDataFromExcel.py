import pandas as pd

sourceonedata = 'E:\Priya\Python\Practice\Citi_Industrialization\CP_AGGR.xlsx'
destdata = 'E:\Priya\Python\Practice\Citi_Industrialization\CP_REVENUE_PACK.xlsx'
resultexcelpath ='E:\Priya\Python\Practice\Citi_Industrialization\Result.xlsx'


def readdatafromexcel(excelpath, sheetname):
    return pd.read_excel(excelpath, sheet_name=sheetname)

def acceptkeyvalues():
    key_input = input("Enter a key element separated by space :")
    return key_input.split()

def acceptresultvalues():
    results_col_input  = input("Enter a result column list separated by space : ")
    return results_col_input.split()

def generatekeycolvalues(df, keylist ,keyname):
    for i, row in df.iterrows():
        keyvalue = ""
        for x in keylist:
            keyvalue = keyvalue + str(row[x])
            print(keyvalue.replace(" ", ""))
        df.at[i, keyname ] = keyvalue.replace(" ", "")



def writedatatoexcel(resultset):
    writer = pd.ExcelWriter(resultexcelpath, engine='xlsxwriter')
    resultset.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.save()

def compareresult(df1,df2 ,resultlist):
    resultset = pd.merge(df1, df2, left_on='key1', right_on='key2')
    for i, row in resultset.iterrows():
        for x in resultlist:
            newresultcolname = x + '_diff'
            check_x_colname = newresultcolname.replace("_diff", "_x")
            check_y_colname = newresultcolname.replace("_diff", "_y")
            diff_value = row[check_x_colname] - row[check_y_colname]
            print("value at %d row and colume %s : %d " % (i, newresultcolname, diff_value))
            resultset.at[i, newresultcolname] = diff_value
    return resultset

def main():
    df1 = readdatafromexcel(sourceonedata, 'CP_AGGR')
    df1['key1'] = ""
    print(df1)
    df2 = readdatafromexcel(destdata, 'CP_REVENUE_PACK')
    df2['key2'] = ""
    print(df2)
    keylist1 = acceptkeyvalues()
    keylist2 = acceptkeyvalues()
    generatekeycolvalues(df1, keylist1 , 'key1')
    generatekeycolvalues(df2, keylist2 , 'key2')
    # print(df1)
    # print(df2)
    resultlist = acceptresultvalues()
    resultset = compareresult(df1 , df2 ,resultlist )
    print(resultset)
    writedatatoexcel(resultset)

if __name__ == '__main__':
    main()



