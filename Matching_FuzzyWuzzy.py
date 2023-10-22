# This is the step 3
import pandas as pd
from fuzzywuzzy import process
import numpy as np

# Load the Data_CC_Collection_From_Teams.xlsx file
teams_df = pd.read_excel('Data_CC_Collection_From_Teams_0822.xlsx', sheet_name='字段整理合集：CMS')

# Load the structured data from the CSV
flattened_df = pd.read_csv('structured_data.csv')

# Convert the 'Fourth_Level_Chinese' column to strings and replace NaN with empty strings
teams_df['名称、业务描述（CN)'] = teams_df['名称、业务描述（CN)'].astype(str).fillna('')

# Create an empty list to store the best matches
third_Chinese_list = []
fourth_Chinese_list = []
tag_list = []

# Perform fuzzy matching and find the best match for each value in column D
for value in teams_df['名称、业务描述（CN)']:
    if value != 'nan':  # Skip if the value is 'nan'
        third_Chinese =
        fourth_Chinese = process.extractOne(value, flattened_df['Fourth_Level_Chinese'])
        tag = process.extractOne(value, flattened_df['Label'])
        fourth_Chinese_list.append(fourth_Chinese[0])
        tag_list.append(tag[0])
    else:
        fourth_Chinese_list.append(np.nan)  # Append NaN if the value is 'nan'
        tag_list.append(np.nan)  # Append NaN if the value is 'nan'

# Add the best matches to a new column in the teams_df DataFrame
teams_df['fourth_Chinese'] = fourth_Chinese_list
teams_df['Label'] = tag_list

# Save the updated DataFrame back to the Excel file
teams_df.to_excel('Data_CC_Collection_From_Teams_with_matches.xlsx', index=False)
