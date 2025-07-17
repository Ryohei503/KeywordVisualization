import pandas as pd
import joblib
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import matplotlib
matplotlib.use('Agg')
from pie_chart_util import generate_category_pie_chart
from word_count_util import word_count


def categorize_summaries():
    root = tk.Tk()
    root.withdraw()
    input_excel = filedialog.askopenfilename(
        title="Select Excel file to categorize",
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
    )
    if not input_excel:
        print("No file selected.")
        return None

    df = pd.read_excel(input_excel, sheet_name=0)
    if 'Summary' not in df.columns:
        messagebox.showerror("Error", "The selected Excel file does not contain a 'Summary' column.")
        return None

    df.dropna(subset=['Summary'], inplace=True)

    pipeline = joblib.load("classifier_model.pkl")

    from sentence_transformers import SentenceTransformer
    st_model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
    def predict_category(row):
        try:
            # Encode the summary using SentenceTransformer
            emb = st_model.encode([row['Summary']])  # returns shape (1, dim)
            proba = pipeline.predict_proba(emb)[0]
            predicted_label = pipeline.classes_[proba.argmax()]
            return predicted_label
        except Exception as e:
            messagebox.showerror("Prediction Error", f"Prediction error: {e}")

    df['Predicted_Category'] = df.apply(predict_category, axis=1)

    # Save categorized Excel
    ordered_categories = ["UI", "API", "Database", "Others"]
    base = os.path.splitext(input_excel)[0]
    output_excel = f"{base}_categorized.xlsx"

    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        for cat in ordered_categories:
            df_cat = df[df['Predicted_Category'] == cat]
            if not df_cat.empty:
                sheet_name = str(cat)[:31]
                original_cols = [col for col in df.columns if col not in ['Predicted_Category']]
                df_cat[original_cols].to_excel(writer, sheet_name=sheet_name, index=False, header=True)


    messagebox.showinfo("Categorized Excel Saved", f"Categorized Excel file saved to:\n{output_excel}")
    return output_excel

# Run categorization and pie chart
output_excel = categorize_summaries()
if output_excel is None or not os.path.exists(output_excel):
    raise ValueError(f"Invalid file path returned by categorize_excel_to_excel: {output_excel}")
generate_category_pie_chart(output_excel)
word_count(output_excel)