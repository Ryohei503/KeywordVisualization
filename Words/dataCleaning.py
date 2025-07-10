from Imports import *

def remove_project_names_and_codes(text, project_names):
    # Remove project names with different variations of brackets and spacing
    for project_name in project_names:
        pattern = re.compile(rf'\s*{re.escape(project_name)}\s*[-]*\s*', re.IGNORECASE)
        text = pattern.sub('', text)

    # Remove patterns like "ABC-100", "AHA-100"
    code_pattern = r'\b[A-Z]{3}-\d+\b'  # Matches patterns like "ABC-100"
    text = re.sub(code_pattern, '', text)

    return text.strip()  # Strip any leading/trailing whitespace


def remove_us_names(text):
    # Define the pattern to match "ABC-100", "1-100", "AHA-100", etc.
    pattern = r'\b[A-Za-z]*\d+-\d+\b'  # Matches any letters followed by digits, a hyphen, and more digits.
    
    # Remove matching patterns
    cleaned_text = re.sub(pattern, '', text)
    
    return cleaned_text.strip()  # Strip any leading/trailing whitespace


def remove_leading_hyphens_and_such(text):
    # function to remove leading hyphens or other unwanted characters
    return re.sub(r'^[\s-]+', '', text) 

def remove_parenthese(text):
    # Regular expression to remove text within parentheses, including the parentheses
    return re.sub(r'\(.*?\)', '', text)  # Remove text in parentheses
 
DetectorFactory.seed = 0

def is_japanese(word):
    # Regular expression to check if the word contains Japanese characters
    return re.search(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]', word) is not None

def spelling_correction(text, preserve_words):
    corrected_text = []
    
    # Split the text into words
    words = text.split()
    
    for word in words:
        # Check if the word is in the preserve list
        if word.lower() in preserve_words or is_japanese(word):
            corrected_text.append(word)  # Keep the original word
        else:
            try:
                # Detect the language of the entire text
                language = detect(text)
                # Only correct if the language is English
                if language == 'en':
                    corrected_word = str(TextBlob(word).correct())  # Correct the word
                    corrected_text.append(corrected_word)
                else:
                    corrected_text.append(word)  # Keep the original word if it's not English
            except Exception as e:
                print(f"Error detecting language for '{word}': {e}")
                corrected_text.append(word)  # In case of any detection error, keep the original word

    return ' '.join(corrected_text)  # Join the corrected words back into a string

def load_preserve_words(filename):
    try:
        with open(filename, 'r') as file:
            return {line.strip().lower() for line in file if line.strip()}  # Load words into a set
    except FileNotFoundError:
        print(f"File {filename} not found. Using an empty set for preserve words.")
        return set()
                
# Function to process the CSV
def process_csv():
    root = Tk()
    root.withdraw()

    # Open the file dialog to select a CSV file
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    
    # Print the file path for debugging
    print(f"Selected file path: {file_path}")
    
    # Check if a file was selected
    if not file_path:
        print("No file selected. Exiting the process.")
        return  # Exit the function if no file was selected

    try:
        # Attempt to read the CSV file
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"Error reading the CSV file: {e}")
        return  # Exit the function if there was an error

    project_names = ["[Project name]"]

    # Load preserve words from the text file
    preserve_words = load_preserve_words('preserve_words.txt')

    if 'Summary' in df.columns:
        print(f"Processing the 'Summary' column with {len(df)} entries. This may take some time...")  # Inform user of processing start
        
        # Using progress indication
        for index, row in df.iterrows():
            # Clean the 'Summary' by removing project names, US names, leading hyphens, and text in parentheses
            cleaned_summary = remove_leading_hyphens_and_such(
                remove_us_names(
                    remove_project_names_and_codes(str(row['Summary']), project_names)
                )
            )
            
            # Remove text in parentheses
            cleaned_summary = remove_parenthese(cleaned_summary)

            # Apply spelling correction
            df.at[index, 'Summary'] = spelling_correction(cleaned_summary, preserve_words)

            # Print progress every 20 rows
            if index % 20 == 0:
                print(f"Processed {index} rows...")

        print("Processing complete!")  # Notify user of completion
    else:
        print("Column 'Summary' not found in the CSV.")

    output_file = 'CleanedOutput.csv'
    df.to_csv(output_file, encoding="utf-8-sig", index=False)
    print(f"Cleaned data saved to {output_file}")

process_csv()
