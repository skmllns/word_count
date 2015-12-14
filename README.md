# word_count
Converts Excel file containing student comments into a form usable by Tableau for data visualization.

# Input

.xlsx file, with headers "Class", "Enrollment", "Gender", "Comments"

# SCRIPT

Python script strips comments and uses regular expressions and NLTK parser to extract nouns, adjectives and verbs. 

# Output

CSV file with an added ID attribute, original headers, and a single word from the corresponding comments per row. 
Imported into Tableau and created a word cloud.

For privacy reasons, the input and output files have not been included.

# Dependencies

openpyxl, re, collections, csv, nltk

