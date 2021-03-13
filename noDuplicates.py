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

def create_dict():
    df=pd.read_excel("noDupsTime.xlsx")
    names_list=list(df['Name'])
    dates_list=list(df['Date/Time'])
    custom_dict={}#dictionary of names as keys and al lthe datetime as values
    modified_dict={}
    for name,date_time  in zip(names_list,dates_list):
        if name not in custom_dict.keys():
            custom_dict[name]=[date_time]
        else:
            custom_dict[name].append(date_time)
    for name in custom_dict.keys():
        #for each name go through each date,then split date_time into date and time,
        #create a dictionary with each day as the key and the values an array of times
        date_dict={}
        for dateTime in custom_dict[name]:
        #iterate over datetimes of each person
            date = dateTime.split()[0]
            time = dateTime.split()[-1]
            if date not in date_dict.keys():
                date_dict[date]=[time]
            else:
                date_dict[date].append(time)
        modified_dict[name]=date_dict#create new dictionary with the name as the key
    print(modified_dict)
    return modified_dict
def create_report(create_dict):
    
    data=create_dict()
    lst_of_names=[]
    length=0
    lst_of_dates=[]
    lst_of_timein=[]
    lst_of_timeout=[]
    lst_of_durations=[]
    total=0
    #iterate over all the names
    for name in data.keys():
        length+=1
        days=data[name]
        dates=data[name].keys()#gets dates for each name 
        print(f'each_day:{type(days)}')
        no_of_names=len(dates)
        lst_of_names.extend([name for i in range(no_of_names)])#make name array same size as dates array
        lst_of_dates.extend(dates)
        for day in days.values(): 
            print('day:',day)
            total+=1
            lst_of_timein.append(day[0])
            lst_of_timeout.append(day[-1])
        #lst_of_durations.extend(data[name].values()[0])
        
    print(f'total:{total}')
    print(f"old length:{length}\nnew length:{len(lst_of_names)}")
    print(f"no of dates:{len(lst_of_dates)}")
    print(f'{len(lst_of_timein)}')
    print(f'{len(lst_of_timeout)}')
    df=pd.DataFrame({"Names":lst_of_names,"Date":lst_of_dates,"Time_in":lst_of_timein,"Time_out":lst_of_timeout})
    #df2=pd.DataFrame({"Time_in":lst_of_timein,"Time_out":lst_of_timeout})
    #df=pd.DataFrame({"Names":lst_of_names,"Date":lst_of_dates})
    df.to_excel('fingers.xlsx')
    #df2.to_excel('fingers2.xlsx')
create_report(create_dict)
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
