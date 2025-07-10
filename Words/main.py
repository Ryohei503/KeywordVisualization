from Imports import *

import Words.dataCleaning as dataCleaning

try:
    stopwords_path = nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

# Load English stopwords
stop_words = set(stopwords.words('english'))



# Read the CSV file
file_path = 'CleanedOutput.csv'
d = dict()
df = pd.read_csv(file_path, encoding='utf-8')

def load_stopwords(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        return {line.strip() for line in file if line.strip()}

# Load the stopwords
japanese_stopwords = load_stopwords('japaneseStopwords.txt')

for index, row in df.iterrows():
    sentence = row['Summary'].strip()  # Add "Summary" on the first row
    sentence = sentence.lower()
    sentence = re.sub(r'[^\w\s]', '', sentence)  # Removes punctuation
    words = re.split(r'[\s]+', sentence)  # Split by any spaces

    # Adding words to the dictionary and checking if words are there
    for word in words:
        if word not in stop_words and japanese_stopwords and len(word) > 1:  # Not stop words
            if word in d:
                d[word] += 1
            else:
                d[word] = 1

# Sort the dictionary by value (count) in descending order
sorted_d = dict(sorted(d.items(), key=lambda item: item[1], reverse=True))

# Print the sorted dictionary and save to CSV and Excel
df_sorted = pd.DataFrame(list(sorted_d.items()), columns=['word', 'count'])
df_sorted.to_csv('SortedKeywords.csv', encoding="utf-8-sig", index=False)
df_sorted.to_excel('ExcelSortedkeywords.xlsx', sheet_name='keywordCount', index=False)

# Print the sorted word counts
for key in sorted_d:
    print(key, ":", sorted_d[key])
