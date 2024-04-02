import pandas as pd

# Define the expected keys
expected_keys = ["Name", "Group", "IP Address", "Login Time", "Logged In", "Idle Time"]

# Initialize a list to hold each record's data
data = []

# Temporary dictionary to hold the current record's data
current_record = {}

# Path to the text file
file_path = 'C:\\Users\\lucas.lee\\Downloads\\sonicwall\\status.txt'

with open(file_path, 'r') as file:
    for line in file:
        # Check if the line contains any of the expected keys
        if any(key + ":" in line for key in expected_keys):
            # Extract the key and value from the line
            key, value = line.split(":", 1)
            key = key.strip()
            value = value.strip()

            # Add the key-value pair to the current record
            current_record[key] = value

            # If the current record has all the expected keys, add it to the data list and reset the record
            if all(key in current_record for key in expected_keys):
                data.append(current_record)
                current_record = {}

# Convert the list of records into a DataFrame
df = pd.DataFrame(data)

# Ensure the DataFrame columns are in the desired order
df = df[expected_keys]

# Save the DataFrame to an Excel file
excel_path = 'C:\\Users\\lucas.lee\\Downloads\\sonicwall\\sonic.xlsx'
df.to_excel(excel_path, index=False, engine='openpyxl')

print(f"Excel file has been created at {excel_path}")
