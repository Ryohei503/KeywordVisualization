import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, cross_val_score, StratifiedKFold, GridSearchCV
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report
from sklearn.utils.class_weight import compute_class_weight
from imblearn.over_sampling import SMOTE
from imblearn.pipeline import Pipeline
import joblib
# Add tkinter for file dialog
import tkinter as tk
from tkinter import filedialog
from sentence_transformers import SentenceTransformer




def load_data(filepath):
    df = pd.read_excel(filepath)
    # Check for required columns
    if 'Summary' not in df.columns or 'Category' not in df.columns:
        root = tk.Tk()
        root.withdraw()
        from tkinter import messagebox
        messagebox.showwarning(
            "Missing Columns",
            "The selected file does not contain both 'Summary' and 'Category' columns."
        )
        exit()
    df.dropna(subset=['Summary', 'Category'], inplace=True)
    return df

# Ask user to select an Excel file
def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel file to train the model",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    return file_path

if __name__ == "__main__":
    file_path = select_excel_file()
    if not file_path:
        print("No file selected. Exiting.")
        exit()
    df = load_data(file_path)
    y = df['Category']

    # Use SentenceTransformer for multilingual embeddings
    model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
    X = model.encode(df['Summary'].astype(str).tolist())  # X is a dense numpy array

    # Train-test split
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=None, stratify=y
    )

    # Calculate class weights for imbalance handling
    classes = np.unique(y_train)
    weights = compute_class_weight('balanced', classes=classes, y=y_train)
    class_weights = dict(zip(classes, weights))

    # Model pipeline with SMOTE and LogisticRegression (no feature engineering needed)
    pipeline = Pipeline([
        ('smote', SMOTE(random_state=42, k_neighbors=5)),
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

    # Save model
    joblib.dump(pipeline, "classifier_model.pkl")
    print("Model saved as 'classifier_model.pkl'")