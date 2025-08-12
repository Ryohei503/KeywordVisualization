import pandas as pd
import numpy as np
import os
## Removed invalid joblib start method for threading
os.environ["LOKY_MAX_CPU_COUNT"] = "1"  # stops loky from launching processes
os.environ["OMP_NUM_THREADS"] = str(os.cpu_count())  # still use all cores for BLAS
from sklearn.model_selection import cross_val_score, StratifiedKFold, GridSearchCV
from sklearn.linear_model import LogisticRegression
from sklearn.utils.class_weight import compute_class_weight
from imblearn.over_sampling import SMOTE
from imblearn.pipeline import Pipeline
import joblib
from sentence_transformers import SentenceTransformer
import os


def train_model(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = [col.lower() for col in df.columns]

        if 'summary' not in df.columns or 'category' not in df.columns:
            return False, "The selected file does not contain both 'Summary' and 'Category' columns."

        df.dropna(subset=['summary', 'category'], inplace=True)
        y = df['category']

        model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
        X = model.encode(df['summary'].astype(str).tolist())

        classes = np.unique(y)
        weights = compute_class_weight('balanced', classes=classes, y=y)
        class_weights = dict(zip(classes, weights))

        pipeline = Pipeline([
            ('smote', SMOTE(random_state=42, k_neighbors=5)),
            ('clf', LogisticRegression(
                class_weight=class_weights,
                max_iter=1000,
                solver='saga',
                n_jobs=-1,  # allow parallel threads
                random_state=42
            ))
        ])

        param_grid = {'clf__C': [0.1, 1, 10]}

        # Use loky backend for joblib (default, supports multiprocessing)
        from joblib import parallel_backend
        with parallel_backend('loky'):
            grid = GridSearchCV(pipeline, param_grid, cv=3, scoring='f1_weighted', n_jobs=-1)
            grid.fit(X, y)
            pipeline = grid.best_estimator_

            cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
            cv_scores = cross_val_score(pipeline, X, y, cv=cv, scoring='f1_weighted', n_jobs=-1)

        print(f"Cross-Validation F1 Score: {cv_scores.mean():.2f} (+/- {cv_scores.std():.2f})")

        # Ask user for output path
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        default_model_name = f"{base_name}_classifier_model.pkl"
        output_path = filedialog.asksaveasfilename(
            title="Save Model As",
            initialfile=default_model_name,
            defaultextension=".pkl",
            filetypes=[("Pickle files", "*.pkl"), ("All files", "*.*")]
        )
        if not output_path:
            return False, "Model save cancelled by user."
        joblib.dump(pipeline, output_path)

        return True, f"Model saved as '{output_path}'"

    except Exception as e:
        return False, str(e)
