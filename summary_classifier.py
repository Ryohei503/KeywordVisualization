import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, cross_val_score, StratifiedKFold, GridSearchCV
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report, confusion_matrix
from sklearn.utils.class_weight import compute_class_weight
from imblearn.over_sampling import SMOTE
from imblearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import StandardScaler
import joblib
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend to avoid Tkinter errors
import matplotlib.pyplot as plt
import seaborn as sns
import re
import string
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer
from nltk.corpus import stopwords
from sklearn.metrics import accuracy_score
import nltk

# warnings.filterwarnings("ignore")

for resource in ['punkt', 'wordnet', 'stopwords']:
    try:
        nltk.data.find(f'tokenizers/{resource}' if resource == 'punkt' else f'corpora/{resource}')
    except LookupError:
        nltk.download(resource, quiet=True)

# Download NLTK resources (run once)
import tkinter as tk
from tkinter import filedialog
import os
from sklearn.metrics import accuracy_score
import csv

def categorize_csv_to_excel(threshold):
    """
    Prompts the user to select a CSV file containing a 'Summary' column,
    predicts categories, and writes an Excel file with separate sheets for each category.
    The output Excel is saved in the same folder as the original file, named 'originalname_categorized.xlsx'.
    Adds a column with the model's confidence for each prediction.
    Adds a column with 1 if the prediction is correct (if true Category exists), else 0.
    Returns the path of the saved Excel file.
    """

    # Open file dialog to select CSV
    root = tk.Tk()
    root.withdraw()
    input_csv = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not input_csv:
        print("No file selected.")
        return None

    # Read the CSV file
    df = pd.read_csv(input_csv, encoding='utf-8-sig')
    if 'Summary' not in df.columns:
        from tkinter import messagebox
        messagebox.showerror("Error", "The selected CSV file does not contain a 'Summary' column.")
        return None

    # Check if the selected file has a 'Category' column
    if 'Category' in df.columns:
        has_category = True
    else:
        has_category = False

    preprocessor = TextPreprocessor()
    df['processed_summary'] = df['Summary'].apply(preprocessor.preprocess)
    df['summary_length'] = df['processed_summary'].str.split().apply(len)
    df['char_count'] = df['processed_summary'].str.len()
    df['digit_count'] = df['processed_summary'].str.count(r'\d')

    def predict_with_confidence(row):
        try:
            input_df = pd.DataFrame([{
                'processed_summary': row['processed_summary'],
                'summary_length': row['summary_length'],
                'char_count': row['char_count'],
                'digit_count': row['digit_count']
            }])
            proba = pipeline.predict_proba(input_df)[0]
            max_prob = max(proba)
            predicted_label = pipeline.classes_[proba.argmax()]
            if max_prob < threshold:
                fallback_label = fallback_predict(row['Summary'])
                pred = fallback_label
            else:
                pred = predicted_label
            # If true label exists, set correct=1 if prediction matches, else 0
            if has_category:
                correct = int(pred == row['Category'])
                return pd.Series([pred, max_prob, correct])
            else:
                return pd.Series([pred, max_prob])
        except Exception as e:
            print(f"Prediction error: {e}")
            fallback_label = fallback_predict(row['Summary'])
            if has_category:
                correct = int(fallback_label == row.get('Category', None))
                return pd.Series([fallback_label, 0.0, correct])
            else:
                return pd.Series([fallback_label, 0.0])

    if has_category:
        df[['Predicted_Category', 'Confidence', 'Correct']] = df.apply(predict_with_confidence, axis=1)
    else:
        df[['Predicted_Category', 'Confidence']] = df.apply(predict_with_confidence, axis=1)

    # If the CSV has a true Category column, print accuracy
    if has_category:
        accuracy = accuracy_score(df['Category'], df['Predicted_Category'])
        print(f"Categorization Accuracy: {accuracy:.2f}")

    # Desired sheet order
    ordered_categories = ["UI", "API", "Database", "Others"]

    # Prepare output path
    base = os.path.splitext(input_csv)[0]
    output_excel = f"{base}_categorized.xlsx"

    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        for cat in ordered_categories:
            df_cat = df[df['Predicted_Category'] == cat]
            if not df_cat.empty:
                # Sheet names in Excel have a max length of 31
                sheet_name = str(cat)[:31]
                # Write columns, no header
                if has_category:
                    df_cat[['Summary', 'Confidence', 'Correct']].to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                else:
                    df_cat[['Summary', 'Confidence']].to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    
    print(f"Excel file saved to {output_excel}")
    return output_excel

# Enhanced text preprocessing
class TextPreprocessor:
    def __init__(self):
        self.lemmatizer = WordNetLemmatizer()
        self.stop_words = set(stopwords.words('english'))
        self.japanese_keywords = {
            '保存', 'フォント', 'ページ', '検索', '戻る', '次へ', 
            '情報', 'キャンセル', '送信', 'ヘルプ', '認証', 'ログイン',
            '更新日時', 'ユーザー一覧', '設定', '注文履歴', '顧客',
            '在庫', '従業員', 'パスワード', 'レポート', 'ダッシュボード',
            '翻訳', 'バックアップ', '通知', '監査', 'ドキュメント'
        }
        
    def preprocess(self, text):
        # Lowercase
        text = text.lower()
        
        # Remove special characters but preserve Japanese text
        text = re.sub(r'[^\w\s\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\uff00-\uff9f\u4e00-\u9faf]', ' ', text)
        
        # Tokenize
        try:
            tokens = word_tokenize(text, language='english')
        except Exception:
            # Fallback: regex-based tokenizer (splits on whitespace)
            tokens = re.findall(r'\w+|[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\uff00-\uff9f\u4e00-\u9faf]+', text)
        
        # Lemmatize and remove stopwords (for English)
        processed_tokens = []
        for token in tokens:
            if token in self.japanese_keywords:
                processed_tokens.append(token)
            else:
                if token not in self.stop_words and token not in string.punctuation:
                    lemma = self.lemmatizer.lemmatize(token)
                    processed_tokens.append(lemma)
        
        return ' '.join(processed_tokens)

# Load and preprocess data
def load_data(filepath):
    df = pd.read_csv(filepath, encoding='utf-8-sig')
    df.dropna(subset=['Summary', 'Category'], inplace=True)
    preprocessor = TextPreprocessor()
    df['processed_summary'] = df['Summary'].apply(preprocessor.preprocess)
    df['summary_length'] = df['processed_summary'].str.split().apply(len)
    df['char_count'] = df['processed_summary'].str.len()
    df['digit_count'] = df['processed_summary'].str.count(r'\d')
    df = df[df['summary_length'] > 3]
    return df

# Load data
df = load_data("sample_train.csv")
X = df[['processed_summary', 'summary_length', 'char_count', 'digit_count']]
y = df['Category']

# Train-test split
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42, stratify=y
)

# Calculate class weights for imbalance handling
classes = np.unique(y_train)
weights = compute_class_weight('balanced', classes=classes, y=y_train)
class_weights = dict(zip(classes, weights))

# Feature processing
text_features = 'processed_summary'
numeric_features = ['summary_length', 'char_count', 'digit_count']

preprocessor = ColumnTransformer([
    ('tfidf', TfidfVectorizer(
        stop_words='english',
        max_features=8000,
        ngram_range=(1, 3),
        min_df=1,
        max_df=0.98
    ), text_features),
    ('num', StandardScaler(), numeric_features)
])

# Model pipeline with SMOTE and LogisticRegression
pipeline = Pipeline([
    ('features', preprocessor),
    ('smote', SMOTE(random_state=42, k_neighbors=5)),  # Set k_neighbors=1 or 2 when there's less than 6 samples in each class
    ('clf', LogisticRegression(
        class_weight=class_weights,
        max_iter=1000,
        solver='saga',
        n_jobs=-1,
        random_state=42
    ))
])

# Hyperparameter tuning
param_grid = {
    'features__tfidf__max_features': [5000, 8000, 12000],
    'clf__C': [0.1, 1, 10]
}
grid = GridSearchCV(pipeline, param_grid, cv=3, scoring='f1_weighted', n_jobs=-1)
grid.fit(X_train, y_train)
pipeline = grid.best_estimator_

# Cross-validation
cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
cv_scores = cross_val_score(pipeline, X_train, y_train, cv=cv, scoring='f1_weighted')
print(f"Cross-Validation F1 Score: {cv_scores.mean():.2f} (+/- {cv_scores.std():.2f})")

# Evaluation
y_pred = pipeline.predict(X_test)
print("\nClassification Report:")
print(classification_report(y_test, y_pred, zero_division=0))

# Confusion Matrix
labels = sorted(y.unique())
cm = confusion_matrix(y_test, y_pred, labels=labels)
plt.figure(figsize=(10, 8))
sns.heatmap(cm, annot=True, fmt='d', xticklabels=labels, yticklabels=labels)
plt.xlabel('Predicted')
plt.ylabel('Actual')
plt.title('Confusion Matrix')
plt.tight_layout()
plt.savefig("confusion_matrix_enhanced.png")
plt.close()

# Save model
joblib.dump(pipeline, "enhanced_classifier.pkl")

# Enhanced fallback system with domain-specific rules
def fallback_predict(summary):
    summary_lower = summary.lower()
    
    # UI-related keywords
    ui_keywords = {
        "button", "form", "screen", "ui", "page", "mobile", "dropdown", 
        "modal", "alignment", "responsive", "layout", "font", "color",
        "click", "hover", "scroll", "display", "render", "visible",
        "保存", "フォント", "ページ", "検索", "戻る", "次へ", "情報",
        "キャンセル", "送信", "ヘルプ", "menu"
    }
    
    # API-related keywords
    api_keywords = {
        "api", "endpoint", "request", "response", "json", "xml", 
        "status", "code", "header", "authentication", "token",
        "rate limit", "timeout", "payload", "webhook", "rest",
        "認証", "ログイン", "更新日時", "ユーザー一覧"
    }
    
    # Database-related keywords
    db_keywords = {
        "database", "query", "table", "sql", "record", "schema",
        "index", "join", "transaction", "constraint", "trigger",
        "replication", "backup", "migration", "performance",
        "注文履歴", "顧客", "在庫", "従業員", "パスワード", "設定"
    }
    
    # Count matches for each category
    ui_matches = sum(1 for kw in ui_keywords if kw in summary_lower)
    api_matches = sum(1 for kw in api_keywords if kw in summary_lower)
    db_matches = sum(1 for kw in db_keywords if kw in summary_lower)
    
    # Return category with most matches
    matches = {
        "UI": ui_matches,
        "API": api_matches,
        "Database": db_matches
    }
    
    max_category = max(matches.items(), key=lambda x: x[1])
    return max_category[0] if max_category[1] > 0 else "Others"

# # Combined predictor with confidence threshold
# def predict_category(summary, threshold):
#     try:
#         # Preprocess
#         preprocessor = TextPreprocessor()
#         processed_text = preprocessor.preprocess(summary)
#         summary_length = len(processed_text.split())
#         char_count = len(processed_text)
#         digit_count = sum(c.isdigit() for c in processed_text)

#         # Prepare input as DataFrame with correct columns
#         input_df = pd.DataFrame([{
#             'processed_summary': processed_text,
#             'summary_length': summary_length,
#             'char_count': char_count,
#             'digit_count': digit_count
#         }])

#         # Predict probabilities
#         proba = pipeline.predict_proba(input_df)[0]
#         max_prob = max(proba)
#         predicted_label = pipeline.classes_[proba.argmax()]

#         # Use fallback if confidence is low
#         if max_prob < threshold:
#             return fallback_predict(summary)
#         return predicted_label
#     except Exception as e:
#         print(f"Prediction error: {e}")
#         return fallback_predict(summary)

# Call the categorization function

def generate_category_pie_chart(excel_file):
    """
    Generates a pie chart showing the proportion of each category based on the number of rows in each category.
    """
    # Read the Excel file
    df = pd.read_excel(excel_file, sheet_name=None, header=None)

    # Count occurrences of each category based on the sheet names
    category_counts = {sheet_name: len(sheet_data) for sheet_name, sheet_data in df.items() if not sheet_data.empty}

    # Create pie chart using seaborn
    labels = list(category_counts.keys())[::-1]  # Reverse the order of categories
    sizes = list(category_counts.values())[::-1]  # Reverse the order of sizes

    # Create a figure and axis
    fig, ax = plt.subplots(figsize=(6, 6))

    # Create the pie chart with the specified style
    patches, texts, pcts = ax.pie(
        sizes, labels=labels, autopct='%.1f%%',
        wedgeprops={'linewidth': 3.0, 'edgecolor': 'white'},
        textprops={'size': 'x-large'},
        startangle=90
    )

    # Set the text label colors to match the wedge face colors
    for i, patch in enumerate(patches):
        texts[i].set_color(patch.get_facecolor())

    # Set properties for percentage text
    plt.setp(pcts, color='black')
    plt.setp(texts, fontweight=600)

    # Set the title
    ax.set_title('Defect Categories', fontsize=18)

    # Adjust layout and save the chart
    plt.tight_layout()
    output_pie_chart = f"{os.path.splitext(excel_file)[0]}_pie_chart.png"
    plt.savefig(output_pie_chart)
    plt.close()
    print(f"Pie chart saved as '{output_pie_chart}'.")

# Call the categorization function
output_excel = categorize_csv_to_excel(0.8)

# Validate the output
if output_excel is None or not os.path.exists(output_excel):
    raise ValueError(f"Invalid file path returned by categorize_csv_to_excel: {output_excel}")

# Generate pie chart after categorization
generate_category_pie_chart(output_excel)