import pandas as pd
from glob import glob


def removeDups():
    df = pd.read_excel("Feb 2021.xlsx")
    # Keep only FIRST record from set of duplicates
    df_first_record = df.drop_duplicates(subset="Date/Time", keep="first")
    #creates an excel file with sorted times
    if glob("noDupsTime.xlsx"):
        pass
    else:
        df_first_record.to_excel("noDupsTime.xlsx", index=False)

# removeDups()

#splits the workbook to two workbooks(with the firt occurence and with the last occurence of a date)
def splitWB():
    dfNoDups = pd.read_excel("test.xlsx")
    df_2_records1 = dfNoDups.drop_duplicates(subset="Date", keep="first")
    df_2_records1.to_excel("firstDate.xlsx", index=False)
    df_2_records2 = dfNoDups.drop_duplicates(subset="Date", keep="last")
    df_2_records2.to_excel("lastDate.xlsx", index=False)


#combines 2 excel files based on their header values
def combine():

    excel1 = 'firstDate.xlsx'
    excel2 = 'lastDate.xlsx'

    df1 = pd.read_excel(excel1)
    df2 = pd.read_excel(excel2)

    values1 = df1['Num', 'Department', 'Name', 'ID', 'Date/Time', 'Date', 'Time']
    values2 = df2['Num', 'Department', 'Name', 'ID', 'Date/Time', 'Date', 'Time']

    dataframes = [values1, values2]

    join = pd.concat(dataframes)

    join.to_excel("output.xlsx")

# combine()


#splits a xlsx file into sheets in the same workbook
def sendtosheet():
    df = pd.read_excel('noDupsTime.xlsx')
    cols = df["Name"].unique()
    # copyfile(file, newfile)

    newfile = "nameSplit2.xlsx"
    writer = pd.ExcelWriter(newfile, engine='openpyxl')
    for myname in cols:
        mydf = df.loc[df["Name"] == myname]
        mydf.to_excel(writer, sheet_name=myname, index=False)
    writer.save()

# sendtosheet()