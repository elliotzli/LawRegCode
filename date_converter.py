import pandas as pd
from datetime import datetime, timedelta

def adjust_dates(dates):
    new_dates = []
    for date in dates:
        if pd.isna(date):
            new_dates.append(pd.NaT)
            continue
        day_of_week, old_date = date.split(", ")
        old_date = datetime.strptime(old_date, "%m/%d/%y")
        new_date = old_date.replace(year=old_date.year + 2)
        
        while new_date.weekday() != old_date.weekday():
            new_date += timedelta(days=1)
        
        new_dates.append(new_date.strftime(f"{day_of_week}, %m/%d/%y"))
    
    return new_dates

df = pd.read_excel("Book1.xlsx")
df['Adjusted Dates'] = adjust_dates(df['2022_Date'])
df.to_excel("Book1.xlsx", index=False)
print("Dates have been adjusted and saved in the new column 'Adjusted Dates'.")