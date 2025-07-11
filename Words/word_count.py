import pandas as pd
import nltk
import re
from nltk.corpus import stopwords
from tkinter import filedialog

# Ensure stopwords are available
try:
    stopwords_path = nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

# Load English stopwords
stop_words = set(stopwords.words('english'))

# Load Japanese stopwords
def load_stopwords(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        return {line.strip() for line in file if line.strip()}

# Select the Excel file
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
japanese_stopwords = load_stopwords('japaneseStopwords.txt')

# Load the Excel file to get sheet names dynamically
xls = pd.ExcelFile(file_path)
target_sheets = xls.sheet_names  # Get all sheet names

# Create an Excel writer to save results
output_file_path = file_path.replace('.xlsx', '_wordcount.xlsx')
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # Process each sheet
    for sheet_name in target_sheets:
        try:
            # Read the specific sheet from the Excel file
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Initialize dictionary for keyword counts
            word_count = {}

            # Process each row in the DataFrame
            for index, row in df.iterrows():
                if 'Summary' in row and pd.notna(row['Summary']):  # Ensure 'Summary' exists and is not NaN
                    sentence = row['Summary'].strip()
                    sentence = sentence.lower()
                    sentence = re.sub(r'[^\w\s]', '', sentence)  # Removes punctuation
                    words = re.split(r'[\s]+', sentence)  # Split by any spaces

                    # Count words
                    for word in words:
                        if word not in stop_words and word not in japanese_stopwords and len(word) > 1:  # Not stop words
                            word_count[word] = word_count.get(word, 0) + 1

            # Sort the dictionary by value (count) in descending order
            sorted_word_count = dict(sorted(word_count.items(), key=lambda item: item[1], reverse=True))

            # Convert the sorted dictionary to a DataFrame
            df_sorted = pd.DataFrame(list(sorted_word_count.items()), columns=['word', 'count'])

            # Save the results to a new sheet in the output file
            df_sorted.to_excel(writer, sheet_name=sheet_name, index=False)

        except Exception as e:
            print(f"An error occurred while processing sheet '{sheet_name}': {e}")

print(f"Excel file with words and their counts have been saved to {output_file_path}")