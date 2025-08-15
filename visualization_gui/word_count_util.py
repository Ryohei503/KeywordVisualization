import pandas as pd
from collections import Counter
from tkinter import messagebox
import os
from text_preprocessing import process_text


def word_count(input_excel):
    output_excel = input_excel.replace(".xlsx", "_wordcount.xlsx")
    xls = pd.ExcelFile(input_excel)
    wrote_any = False
    # Use a temporary writer, only create file if data is written
    error_message = None
    try:
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                df.columns = [col.lower() for col in df.columns]
                if 'summary' not in df.columns:
                    error_message = "Summary' column not found."
                    break
                all_tokens = []
                for tokens_str in process_text(df['summary']):
                    all_tokens.extend(str(tokens_str).split())
                if all_tokens:
                    word_counts = Counter(all_tokens)
                    word_count_df = pd.DataFrame({'Word': list(word_counts.keys()), 'Count': list(word_counts.values())})
                    word_count_df.sort_values('Count', ascending=False).to_excel(writer, sheet_name=sheet_name[:31], index=False)
                    wrote_any = True
        # Writer is closed here
        if error_message:
            if os.path.exists(output_excel):
                os.remove(output_excel)
            return None, False, error_message
        if wrote_any:
            return output_excel, True, None
        else:
            if os.path.exists(output_excel):
                os.remove(output_excel)
            return None, False, "No word count data was generated."
    except Exception as e:
        try:
            if os.path.exists(output_excel):
                os.remove(output_excel)
        except Exception:
            pass
        raise