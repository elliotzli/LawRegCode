import pandas as pd
from datetime import datetime

date_today = datetime.now().strftime('%m-%d-%Y')

file_path1 = '//Users/elliotli/Library/CloudStorage/Box-Box/trigger word daily update/00keywordhits take-home old.csv'
file_path2 = '/Users/elliotli/Library/CloudStorage/Box-Box/1_WPDOCS/1 - Exams & Grades Team/Exams/Exam Administrator/Original Exams/24A Exams/Final Exams/Take-home Exams/Faculty Exam Submissions/00keywordhits.csv'

df1 = pd.read_csv(file_path1)
df2 = pd.read_csv(file_path2)

df_updates = df2.merge(df1, indicator=True, how='outer').query('_merge == "left_only"').drop('_merge', axis=1)

df_updates.to_csv(f'take-home exam daily updates {date_today}.csv', index=False)

file_path3 = '/Users/elliotli/Library/CloudStorage/Box-Box/trigger word daily update/00keywordhits in-class old.csv'
file_path4 = '/Users/elliotli/Library/CloudStorage/Box-Box/1_WPDOCS/1 - Exams & Grades Team/Exams/Exam Administrator/Original Exams/24A Exams/Final Exams/In-Class Exams/Faculty Exam Submissions/00keywordhits.csv'

df3 = pd.read_csv(file_path3)
df4 = pd.read_csv(file_path4)

df_updates2 = df4.merge(df3, indicator=True, how='outer').query('_merge == "left_only"').drop('_merge', axis=1)

df_updates2.to_csv(f'in-class exam daily updates {date_today}.csv', index=False)
