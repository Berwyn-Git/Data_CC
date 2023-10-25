# Data_CC
This is an NLP project that recognizes and learns the fields in the business that need to be classified for data classification, and automatically performs the classification according to the classification document guideline requirements.

# How to use
1. Run the `Extracting_Data.py`. This is to get all the labels and their structures from the Excel file.
2. Run the `Chinese_Filling.py`.
3. Run the `Matching_FuzzyWuzzy.py`

# Challenges
## Accuracy
Now the project does not work well on the understanding of the words meaning and it will return the wrong answers.

Solution:
1. Take the English meaning into account. In the file there are English meaning also, so may be that will help on understanding the meaning of all the words.
2. Change the logic of the words embedding. Reset the function of calculating, to focus the attention on the true important words.

Add the current challenges and my thinking of solutions.
