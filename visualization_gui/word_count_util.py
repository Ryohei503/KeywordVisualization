import pandas as pd
from collections import Counter
from tkinter import messagebox
import os
from text_preprocessing import process_text


def word_count(input_excel):
    xls = pd.ExcelFile(input_excel)
    output_excel = input_excel.replace(".xlsx", "_wordcount.xlsx")
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
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

    messagebox.showinfo("Word Count Saved", f"Word count Excel file saved to:\n{output_excel}")
    os.startfile(output_excel)
