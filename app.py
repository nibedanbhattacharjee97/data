import pandas as pd

# Load the two Excel files
project_age_limit = pd.read_excel('project_age_limit.xlsx')
students_data = pd.read_excel('students_data.xlsx')

# Load the Excel file containing already rejected students (with only 'Enrollment ID')
rejected_students = pd.read_excel('rejected_students.xlsx')  # Adjust the file name as needed

# Merge the two dataframes based on 'Concatenate'
merged_data = pd.merge(students_data, project_age_limit, on='Concatenate', how='left')

# Mark the students who are already rejected based on 'Enrollment ID'
merged_data['Already Rejected'] = merged_data['Enrollment ID'].isin(rejected_students['Enrollment ID'])

# Define a function to check if the age is within the allowed range
def check_age(row):
    if row['Already Rejected']:  # If already rejected, skip the age check
        return 'Already Rejected'
    elif row['ACTUAL AGE'] < row['MIN Age']:
        return 'Rejected: Below Minimum Age'
    elif row['ACTUAL AGE'] > row['Max Age']:
        return 'Rejected: Above Maximum Age'
    else:
        return 'Accepted'

# Apply the function to each row and create a new column 'Status'
merged_data['Status'] = merged_data.apply(check_age, axis=1)

# Select only 'Enrollment ID' and 'Status' columns
output_data = merged_data[['Enrollment ID', 'Status']]

# Remove duplicates based on 'Enrollment ID'
output_data = output_data.drop_duplicates(subset=['Enrollment ID'])

# Save the result to a new Excel file
output_data.to_excel('student_status.xlsx', index=False)

print("Process completed. The output file with Enrollment ID and Status is saved as 'student_status.xlsx'.")
