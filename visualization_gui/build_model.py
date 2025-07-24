import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, cross_val_score, StratifiedKFold, GridSearchCV
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report
from sklearn.utils.class_weight import compute_class_weight
from imblearn.over_sampling import SMOTE
from imblearn.pipeline import Pipeline
import joblib
from sentence_transformers import SentenceTransformer
from tkinter import messagebox
import os



def train_model(file_path):
    df = pd.read_excel(file_path)
    # Normalize column names to lowercase
    df.columns = [col.lower() for col in df.columns]
    # Check for required columns
    if 'summary' not in df.columns or 'category' not in df.columns:
        messagebox.showerror(
            "Missing Columns",
            "The selected file does not contain both 'Summary' and 'Category' columns."
        )
        return None
    # Drop rows with missing values in 'summary' or 'category'
    df.dropna(subset=['summary', 'category'], inplace=True)

    y = df['category']

    # Use SentenceTransformer for multilingual embeddings
    model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
    X = model.encode(df['summary'].astype(str).tolist())  # X is a dense numpy array

    # Train-test split
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
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

    # Save model with filename based on input file
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_name = f"{base_name}_classifier_model.pkl"
    joblib.dump(pipeline, output_name)
    messagebox.showinfo("Model Saved", f"Model saved as '{output_name}'")
    