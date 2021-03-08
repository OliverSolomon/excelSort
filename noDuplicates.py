import pandas as pd

df = pd.read_excel("Feb 2021.xlsx")

# Keep only FIRST record from set of duplicates
df_first_record = df.drop_duplicates(subset="Date/Time", keep="first")
df_first_record.to_excel("noDupsTime.xlsx", index=False)
