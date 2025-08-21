import pandas as pd
import joblib
from tkinter import messagebox
import matplotlib
matplotlib.use('Agg')
from sentence_transformers import SentenceTransformer


def categorize_summaries(input_excel, model_path=None, should_cancel=None):
    try:
        df = pd.read_excel(input_excel, sheet_name=0)
        # Normalize column names to lowercase
        df.columns = [col.lower() for col in df.columns]
        if 'summary' not in df.columns:
            messagebox.showerror("Error", "The selected Excel file does not contain a 'Summary' column.")
            return False

        df.dropna(subset=['summary'], inplace=True)

        # model_path is now provided by the GUI, do not prompt again
        if not model_path:
            messagebox.showinfo("No Selection", "No model file selected for categorization.")
            return False
        pipeline = joblib.load(model_path)
        st_model = SentenceTransformer('paraphrase-multilingual-mpnet-base-v2')
        
        predicted_categories = []
        for _, row in df.iterrows():
            # Check for cancellation between each prediction
            if should_cancel and should_cancel():
                return False
                
            try:
                # Encode the summary using SentenceTransformer
                emb = st_model.encode([row['summary']])  # returns shape (1, dim)
                proba = pipeline.predict_proba(emb)[0]
                predicted_label = pipeline.classes_[proba.argmax()]
                predicted_categories.append(predicted_label)
            except Exception as e:
                messagebox.showerror("Prediction Error", f"Prediction error: {e}")
                predicted_categories.append("Error")

        df['Predicted_Category'] = predicted_categories

        # Check for cancellation before saving
        if should_cancel and should_cancel():
            return False

        # Save categorized Excel with categories ordered alphabetically, but 'Others' last
        categories = sorted([cat for cat in df['Predicted_Category'].unique() if cat != 'Others'])
        ordered_categories = categories + (['Others'] if 'Others' in df['Predicted_Category'].unique() else [])
        output_excel = input_excel.replace(".xlsx", "_categorized.xlsx")

        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            for cat in ordered_categories:
                # Check for cancellation between each sheet
                if should_cancel and should_cancel():
                    return False
                    
                df_cat = df[df['Predicted_Category'] == cat]
                if not df_cat.empty:
                    sheet_name = str(cat)[:31]
                    original_cols = [col for col in df.columns if col not in ['Predicted_Category']]
                    df_cat[original_cols].to_excel(writer, sheet_name=sheet_name, index=False, header=True)
        
        return output_excel  # Return the output path instead of showing message

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during categorization:\n{e}")
        return False