import shutil
import os

source_path = '/Users/elliotli/Library/CloudStorage/Box-Box/1_WPDOCS/1 - Exams & Grades Team/Exams/Exam Administrator/Original Exams/24A Exams/Final Exams/Take-home Exams/Faculty Exam Submissions/00keywordhits.csv'
destination_directory = '/Users/elliotli/Library/CloudStorage/Box-Box/trigger word daily update'
new_filename = '00keywordhits take-home old.csv'
destination_path = os.path.join(destination_directory, new_filename)

os.makedirs(destination_directory, exist_ok=True)

try:
    shutil.copy(source_path, destination_path)
    print("File copied and renamed successfully!")
except Exception as e:
    print(f"An error occurred while copying the file: {e}")

source_path = '/Users/elliotli/Library/CloudStorage/Box-Box/1_WPDOCS/1 - Exams & Grades Team/Exams/Exam Administrator/Original Exams/24A Exams/Final Exams/In-Class Exams/Faculty Exam Submissions/00keywordhits.csv'
destination_directory = '/Users/elliotli/Library/CloudStorage/Box-Box/trigger word daily update'
new_filename = '00keywordhits in-class old.csv'
destination_path = os.path.join(destination_directory, new_filename)

os.makedirs(destination_directory, exist_ok=True)

try:
    shutil.copy(source_path, destination_path)
    print("File copied and renamed successfully!")
except Exception as e:
    print(f"An error occurred while copying the file: {e}")




