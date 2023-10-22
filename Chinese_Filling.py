# This is the step 2
import pandas as pd
from fuzzywuzzy import process

# Read the Excel file
df = pd.read_excel('Data_CC_Collection_From_Teams.xlsx', sheet_name='字段整理合集：CMS')

# # Get the column names
# column_names = df.columns
# print(column_names)

# Extract all the field items
filtered_df = df[(df['名称、业务描述（CN)'].isnull()) & (df['业务是否在用？'] == 'Y')]
result_cells = filtered_df['UI Filed name'].tolist()
# print(result_cells)


# Function to find the closest match in column D using fuzzywuzzy
def find_closest_match(value, column_D):
    # match, score = process.extractOne(value, column_D)
    match = process.extractOne(value, column_D)[0]
    score = process.extractOne(value, column_D)[1]
    return match, score


# Loop through result_cells and update corresponding values in Column E
for cell_value in result_cells:
    closest_match, match_score = find_closest_match(cell_value, df['UI Filed name'])

    # Find the index of the closest match in the DataFrame
    match_index = df[df['UI Filed name'] == closest_match].index[0]

    # Update the value in Column E for the current row
    filtered_df.loc[filtered_df['名称、业务描述（CN)'] == cell_value, '名称、业务描述（CN)'] = df.loc[match_index, '名称、业务描述（CN)']

# # Print or use the updated DataFrame
# print(filtered_df)

# Write the outcome into Excel file
output_file = 'Updated_Data.xlsx'
filtered_df.to_excel(output_file, index=False, sheet_name='Updated_Sheet')

print(f"Updated data has been written to {output_file}")
