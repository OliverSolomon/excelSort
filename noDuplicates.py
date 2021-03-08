import pandas as pd
from glob import glob

df = pd.read_excel("Feb 2021.xlsx")

def removeDups():
    # Keep only FIRST record from set of duplicates
    df_first_record = df.drop_duplicates(subset="Date/Time", keep="first")
    #creates an excel file with sorted times
    if glob("noDupsTime.xlsx"):
        pass
    else:
        df_first_record.to_excel("noDupsTime.xlsx", index=False)

# removeDups()

