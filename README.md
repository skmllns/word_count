# word_count
Converts Excel file containing student comments into a form usable by Tableau for data visualization.

# Input

.xlsx file

SAMPLE

|------------------------------------------------------------|
|---class----|-enrollment-|-gender-|-------comments----------|
|------------------------------------------------------------|
| First-year | Full-Time  | Gender | I love this university. |
|------------------------------------------------------------|

# SCRIPT

Python script strips comments and uses regular expressions and NLTK parser to extract nouns, adjectives and verbs. 

# Output

CSV file with an added ID attribute, original headers, and a single word from the corresponding comments per row. 
Imported into Tableau, craeted word cloud which can be found here: https://public.tableau.com/profile/sarah.mullins6785#!/vizhome/NSSEWordCloud/Sheet1

