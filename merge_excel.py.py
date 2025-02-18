import pandas as pd
import os

# Get the current folder path
current_folder = os.path.dirname(os.path.abspath(__file__))

# Find all .xlsx files in the directory
file_paths = [os.path.join(current_folder, file) for file in os.listdir(current_folder) if file.endswith('.xlsx')]
if not file_paths:
    print("No .xlsx files found in the directory.")
    exit()

merged_df = pd.DataFrame()

# Iterate through each Excel file and merge it into the main DataFrame
for file_path in file_paths:
    file_name = os.path.basename(file_path).split('.')[0]  # Extract the file name without extension
    df = pd.read_excel(file_path, sheet_name="Sensor tracing")  # Read only the "Sensor tracing" sheet
    df['Modem'] = file_name  # Add a new column with the file name
    merged_df = pd.concat([merged_df, df], ignore_index=True)

# Remove the first column
merged_df.drop(columns=merged_df.columns[0], inplace=True)

# Reorder columns so that the "Modem" column is first
merged_df = merged_df[['Modem'] + [col for col in merged_df.columns if col != 'Modem']]

# Dictionary mapping values to their replacements
replace_dict = {
    1437256070: 'Beacon1',
    1437262290: 'Beacon10',
    1437254170: 'Beacon2',
    1437241974: 'Beacon3',
    1437253866: 'Beacon4',
    1437253210: 'Beacon5',
    1437238186: 'Beacon6',
    1437242174: 'Beacon7',
    1437255806: 'Beacon8',
    1437252322: 'Beacon9'
}

# Perform replacement using the defined dictionary
merged_df.replace(replace_dict, inplace=True, regex=False)

# Check if the values were replaced correctly
for key, value in replace_dict.items():
    if key in merged_df.values:
        print(f"Value '{key}' was replaced with '{value}'.")
    else:
        print(f"Value '{key}' was not found in the DataFrame.")

# Save the merged DataFrame to a new Excel file
merged_df.to_excel('merged_file.xlsx', index=False)

print("The merged file has been successfully created!")
