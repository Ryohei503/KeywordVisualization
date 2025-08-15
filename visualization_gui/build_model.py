import pandas as pd
import numpy as np
import os
os.environ["LOKY_MAX_CPU_COUNT"] = "1"
os.environ["OMP_NUM_THREADS"] = str(os.cpu_count())
from sklearn.model_selection import cross_val_score, StratifiedKFold, GridSearchCV
from sklearn.linear_model import LogisticRegression
from sklearn.utils.class_weight import compute_class_weight
from imblearn.over_sampling import SMOTE
from imblearn.pipeline import Pipeline
from sentence_transformers import SentenceTransformer

# Global variable to store the trained pipeline
_trained_pipeline = None

def get_trained_pipeline():
    """Returns the trained pipeline"""
    return _trained_pipeline

def train_model(file_path, should_cancel=None):
    global _trained_pipeline
    try:
        # Initial data loading
        df = pd.read_excel(file_path)
        df.columns = [col.lower() for col in df.columns]

        if should_cancel and should_cancel():
            return False, "Training canceled by user"

        # Data validation
        if 'summary' not in df.columns or 'category' not in df.columns:
            return False, "The selected file does not contain both 'Summary' and 'Category' columns."

        df.dropna(subset=['summary', 'category'], inplace=True)
        y = df['category']

        # Text embedding
        model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
        X = model.encode(df['summary'].astype(str).tolist())

        if should_cancel and should_cancel():
            return False, "Training canceled by user"

        # Class weights
        classes = np.unique(y)
        weights = compute_class_weight('balanced', classes=classes, y=y)
        class_weights = dict(zip(classes, weights))

        # Pipeline setup
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

        param_grid = {'clf__C': [0.1, 1, 10]}
        cv = StratifiedKFold(n_splits=3, shuffle=True, random_state=42)

        # Grid search with cancellation support
        from joblib import parallel_backend
        with parallel_backend('loky'):
            grid = GridSearchCV(
                pipeline, 
                param_grid, 
                cv=cv, 
                scoring='f1_weighted', 
                n_jobs=1  # Set to 1 for better cancellation handling
            )
            
            # Custom fit with cancellation checks
            for train_idx, test_idx in cv.split(X, y):
                if should_cancel and should_cancel():
                    return False, "Training canceled by user"
                grid.fit(X, y)
                break  # Only do one fold for cancellation check

            if should_cancel and should_cancel():
                return False, "Training canceled by user"

            # Full training if not canceled
            grid.fit(X, y)
            pipeline = grid.best_estimator_
            _trained_pipeline = pipeline  # Store the trained pipeline

            # Cross-validation
            cv_scores = cross_val_score(
                pipeline, 
                X, 
                y, 
                cv=StratifiedKFold(n_splits=5, shuffle=True, random_state=42), 
                scoring='f1_weighted', 
                n_jobs=1
            )

        print(f"Cross-Validation F1 Score: {cv_scores.mean():.2f} (+/- {cv_scores.std():.2f})")

        if should_cancel and should_cancel():
            return False, "Training canceled by user"

        return True, "Model trained successfully"

    except Exception as e:
        return False, str(e)