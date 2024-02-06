import pandas as pd
import json
import os
import shutil
import os

# Clear the console
os.system('cls' if os.name == 'nt' else 'clear')

# Set console formatting
def print_middle(message, color='white'):
    padding = ' ' * ((shutil.get_terminal_size().columns - len(message)) // 2)
    print(f'{padding}{message}')

# Print Script Title
#################################
padding = '=' * shutil.get_terminal_size().columns
print('\033[91m' + padding, end='')  # Red color
print_middle("MITS - O366 Log Analyzer")
print('\033[31m' + "                                                      version 0.0.7")  # Dark red color
print('\033[91m' + padding, end='')  # Red color
print(" ")
print(" ")

# Reset console output color
print('\033[0m')

# Prompt the user to supply the path and file name to the CSV file
csv_file = input("Please enter the path and file name to the CSV file: ")

# Load the CSV file
df = pd.read_csv(csv_file)

# Check if 'AuditData' column exists
if 'AuditData' in df.columns:
    # Define a function to parse JSON data from 'AuditData'
    def parse_audit_data(json_str):
        try:
            return json.loads(json_str)
        except ValueError as e:
            # Return None or {} if JSON data is invalid
            return None

    # Apply the function to the 'AuditData' column
    df['ParsedAuditData'] = df['AuditData'].apply(parse_audit_data)

    # Now, df['ParsedAuditData'] contains the parsed JSON data
    # You can extract specific incident response data from this column
    items = ["CreationTime", "UserId", "Workload", "Operation", "EventSource", "ClientIP", 
             "AuthenticationType", "SourceFileName", "UserAgent", "SiteUrl", "DeviceDisplayName", 
             "IsManagedDevice"]

    for item in items:
        if 'ParsedAuditData' in df.columns and df['ParsedAuditData'].apply(lambda x: x is not None and item in x).any():
            df[item] = df['ParsedAuditData'].apply(lambda x: x.get(item) if x else None)

    # Create a new DataFrame that only includes the specified columns
    df_filtered = df[items]

    # Save the new DataFrame to a new CSV file
    df_filtered.to_csv('Processed_ActivityResults.csv', index=False)
else:
    print("'AuditData' column not found in the CSV file.")
