import pandas as pd
from collections import Counter
from tkinter import messagebox
import os
from text_preprocessing import process_text


def word_count(excel_path):
    xls = pd.ExcelFile(excel_path)
    base_path = os.path.splitext(excel_path)[0]
    word_count_excel = f"{base_path}_wordcount.xlsx"
    with pd.ExcelWriter(word_count_excel, engine='xlsxwriter') as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            # Normalize column names to lowercase
            df.columns = [col.lower() for col in df.columns]
            if 'summary' not in df.columns:
                messagebox.showerror("Missing Column", f"'Summary' column not found.")
                return None
            all_tokens = []
            for tokens_str in process_text(df['summary']):
                all_tokens.extend(str(tokens_str).split())
            word_counts = Counter(all_tokens)
            word_count_df = pd.DataFrame({'Word': list(word_counts.keys()), 'Count': list(word_counts.values())})
            word_count_df.sort_values('Count', ascending=False).to_excel(writer, sheet_name=sheet_name[:31], index=False)

    messagebox.showinfo("Word Count Saved", f"Word count Excel file saved to:\n{word_count_excel}")
    return word_count_excel
