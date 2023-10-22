import pandas as pd
from fuzzywuzzy import process
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score

# Load the Excel file
df = pd.read_excel('Data_CC_Collection_From_Teams.xlsx', sheet_name='字段整理合集：CMS')

# Extract all the field items
filtered_df = df[(df['名称、业务描述（CN)'].isnull()) & (df['业务是否在用？'] == 'Y')]
result_cells = filtered_df['UI Filed name'].tolist()

# Load the supervised data
supervised_data = pd.read_excel('standard.xlsx', sheet_name='Sheet1')

# Extract relevant columns from supervised data
original_text = supervised_data['Original Text']
best_match = supervised_data['Best Match']
labels = supervised_data['Label']


# Function to find the closest match using fuzzywuzzy
def find_closest_match(value, candidates):
    match, score = process.extractOne(value, candidates)
    return match, score


# Create a supervised dataset
supervised_dataset = []
for cell_value in result_cells:
    closest_match, match_score = find_closest_match(cell_value, original_text)
    match_label = 1 if closest_match in best_match.values.tolist() else 0
    supervised_dataset.append([cell_value, closest_match, match_label])

# Create a DataFrame from the supervised dataset
supervised_df = pd.DataFrame(supervised_dataset, columns=['Input Text', 'Best Match', 'Label'])

# Split the dataset into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(supervised_df[['Input Text', 'Best Match']], supervised_df['Label'],
                                                    test_size=0.2, random_state=42)

# Vectorize the text data using TF-IDF
vectorizer = TfidfVectorizer()
X_train_vectorized = vectorizer.fit_transform(X_train['Input Text'])
X_test_vectorized = vectorizer.transform(X_test['Input Text'])

# Train a machine learning model (Random Forest) on the supervised dataset
model = RandomForestClassifier()
model.fit(X_train_vectorized, y_train)

# Make predictions on the test set
y_pred = model.predict(X_test_vectorized)

# Evaluate the model's accuracy
accuracy = accuracy_score(y_test, y_pred)
print(f"Accuracy: {accuracy:.2f}")

# Now you can use the trained model to predict matches for your original data.
# Loop through result_cells and update corresponding values in Column E
for cell_value in result_cells:
    closest_match, _ = find_closest_match(cell_value, original_text)
    input_text_vectorized = vectorizer.transform([cell_value])
    match_label = model.predict(input_text_vectorized)

    if match_label == 1:
        match_index = best_match[best_match == closest_match].index[0]
        filtered_df.loc[filtered_df['名称、业务描述（CN)'] == cell_value, '名称、业务描述（CN)'] = best_match[match_index]

# Write the outcome into Excel file
output_file = 'Updated_Data.xlsx'
filtered_df.to_excel(output_file, index=False, sheet_name='Updated_Sheet')

print(f"Updated data has been written to {output_file}")
