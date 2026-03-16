# create_excel_files.py
import pandas as pd
import os

# Create data directory
os.makedirs('data', exist_ok=True)

print("Creating Excel files...")

# Student columns
student_cols = ['SR#','Name','Registration No','Room No','Status','Contact No','Father Contact','Blood Group','Semester','Program']
pd.DataFrame(columns=student_cols).to_excel('data/students.xlsx', index=False)
print("✓ Created students.xlsx")

# History columns
history_cols = ['Date','Time','Student Name','Registration No','Room No','Status','Semester','Amount','Payment Method','Remarks']
pd.DataFrame(columns=history_cols).to_excel('data/history.xlsx', index=False)
print("✓ Created history.xlsx")

# Forms columns
forms_cols = ['Student Name','Registration No','Room No','Status','Semester','Admission Form','PWWF Form','Consent Form','Amount']
pd.DataFrame(columns=forms_cols).to_excel('data/forms.xlsx', index=False)
print("✓ Created forms.xlsx")

# Defaulters columns
defaulters_cols = ['Student Name','Registration No','Room No','Status','Semester','Amount','Defaulter Status','Remarks']
pd.DataFrame(columns=defaulters_cols).to_excel('data/defaulters.xlsx', index=False)
print("✓ Created defaulters.xlsx")

# PWWF Boarding columns
pwwf_cols = ['SR#','Student Name','Registration No','Semester','Amount','Paying Date']
pd.DataFrame(columns=pwwf_cols).to_excel('data/pwwf_boarding.xlsx', index=False)
print("✓ Created pwwf_boarding.xlsx")

print("\n✅ All Excel files created successfully in 'data' folder!")
print("Folder location:", os.path.abspath('data'))