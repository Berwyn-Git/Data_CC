# This is the step 1
import pandas as pd

# First should input all the Business-Names, all the classifications, contents and labels
# Read excel file
df_1 = pd.read_excel('Data_CC_Collection.xlsx', sheet_name='字段整理合集：CMS')
df_2 = pd.read_excel('Data_CC_Collection.xlsx', sheet_name='MBFS分类分级')

# Get all the Business-Names
BN = df_1.loc[1:, 'BN']
BN_DF = pd.DataFrame({'Business_Name': BN})
BN_DF.to_csv('business_name.csv', index=False)

# Get all the other data
hierarchy = {}  # Top-level dictionary to hold the hierarchy

for index, row in df_2.iterrows():
    first_level_chinese = row['A']
    first_level_english = row['B']
    second_level_chinese = row['C']
    second_level_english = row['D']
    description_2nd = row['E']
    third_level_chinese = row['F']
    third_level_english = row['G']
    description_3rd = row['H']
    fourth_level_chinese = row['I']
    fourth_level_english = row['J']
    description_4th = row['K']
    label = row['R']

    # Create or update the dictionary structure
    if first_level_chinese not in hierarchy:
        hierarchy[first_level_chinese] = {
            'English': first_level_english,
            'Children': {}
        }
    if second_level_chinese not in hierarchy[first_level_chinese]['Children']:
        hierarchy[first_level_chinese]['Children'][second_level_chinese] = {
            'English': second_level_english,
            'Description': description_2nd,
            'Children': {}
        }
    if third_level_chinese not in hierarchy[first_level_chinese]['Children'][second_level_chinese]['Children']:
        hierarchy[first_level_chinese]['Children'][second_level_chinese]['Children'][third_level_chinese] = {
            'English': third_level_english,
            'Description': description_3rd,
            'Children': {}
        }
    if fourth_level_chinese not in hierarchy[first_level_chinese]['Children'][second_level_chinese]['Children'][third_level_chinese][
                'Children']:
        hierarchy[first_level_chinese]['Children'][second_level_chinese]['Children'][third_level_chinese]['Children'][
            fourth_level_chinese] = {
            'English': fourth_level_english,
            'Description': description_4th,
            'Label': label
        }

# Access first-level data
for first_level, data in hierarchy.items():
    print("First Level Chinese:", first_level)
    print("First Level English:", data['English'])

    # Access second-level data
    for second_level, second_data in data['Children'].items():
        print("Second Level Chinese:", second_level)
        print("Second Level English:", second_data['English'])
        print("Description:", second_data['Description'])

        # Access third-level data
        for third_level, third_data in second_data['Children'].items():
            print("Third Level Chinese:", third_level)
            print("Third Level English:", third_data['English'])
            print("Description:", third_data['Description'])

            # Access fourth-level data
            for fourth_level, fourth_data in third_data['Children'].items():
                print("Fourth Level Chinese:", fourth_level)
                print("Fourth Level English:", fourth_data['English'])
                print("Description:", fourth_data['Description'])
                print("Label:", fourth_data['Label'])

# Flatten the hierarchy and store in a list of dictionaries
flattened_data = []

for first_level, first_data in hierarchy.items():
    for second_level, second_data in first_data['Children'].items():
        for third_level, third_data in second_data['Children'].items():
            for fourth_level, fourth_data in third_data['Children'].items():
                flattened_data.append({
                    'First_Level_Chinese': first_level,
                    'First_Level_English': first_data['English'],
                    'Second_Level_Chinese': second_level,
                    'Second_Level_English': second_data['English'],
                    'Description_2nd': second_data['Description'],
                    'Third_Level_Chinese': third_level,
                    'Third_Level_English': third_data['English'],
                    'Description_3rd': third_data['Description'],
                    'Fourth_Level_Chinese': fourth_level,
                    'Fourth_Level_English': fourth_data['English'],
                    'Description_4th': fourth_data['Description'],
                    'Label': fourth_data['Label']
                })

# Create a DataFrame from the flattened data
flattened_df = pd.DataFrame(flattened_data)

# Save the DataFrame to a CSV file
flattened_df.to_csv('structured_data.csv', index=False)