import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import os
import pandas as pd
from wordcloud import WordCloud
import sys

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# Load bundled Japanese font (ipaexg.ttf)
def load_japanese_font():
    try:
        if getattr(sys, 'frozen', False):
            font_path = os.path.join(sys._MEIPASS, "ipaexg.ttf")
        else:
            font_path = os.path.join(os.path.dirname(__file__), "ipaexg.ttf")
        fm.fontManager.addfont(font_path)
        font_name = fm.FontProperties(fname=font_path).get_name()
        plt.rcParams['font.family'] = font_name
        return font_path
    except Exception as e:
        print(f"Font load failed: {e}")
        return None

# Ask the user to select an Excel file before opening the UI
def select_initial_file():
    file_path = filedialog.askopenfilename(
        title="Excelファイルを選択してください",
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
    )
    if not file_path:
        messagebox.showinfo("キャンセル", "ファイルが選択されませんでした。プログラムを終了します。")
        sys.exit()
    return file_path

# Get sheet names from the selected Excel file
def get_sheet_names(file_path):
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names

# Main program starts here
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    selected_excel_file = select_initial_file()
    sheet_names = get_sheet_names(selected_excel_file)

    root.deiconify()
    root.title("Graph and Word Cloud Generator")
    root.geometry("400x300")

    # Variable to hold the selected sheet name
    selected_sheet = tk.StringVar(root)
    selected_sheet.set(sheet_names[0])  # Set default value

    # Dropdown menu for sheet selection
    sheet_menu = tk.OptionMenu(root, selected_sheet, *sheet_names)
    sheet_menu.pack(pady=20)

    def generate_graph():
        try:
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("エラー", "フォントの読み込みに失敗しました。")
                return
            df = pd.read_excel(selected_excel_file, sheet_name=selected_sheet.get())

            df_sorted = df.sort_values(by="count", ascending=True)
            num_words = len(df_sorted)
            height_per_word = 0.25
            base_height = 4
            fig_height = max(base_height, num_words * height_per_word)

            base_path = os.path.splitext(selected_excel_file)[0]
            sheet_name = selected_sheet.get().replace(" ", "_")  # Replace spaces with underscores for file name
            output_path = f"{base_path}_{sheet_name}_graph.png"
            fig, ax = plt.subplots(figsize=(12, fig_height), constrained_layout=True)
            bars = ax.barh(df_sorted["word"], df_sorted["count"], color="skyblue")
            ax.invert_yaxis()
            ax.set_xlabel("Count", fontsize=14)
            ax.set_ylabel("Word", fontsize=14)
            ax.set_title(f"Word Frequency Graph - {selected_sheet.get()}", fontsize=16)
            ax.set_ylim(-0.5, len(df_sorted) - 0.5)
            ax.tick_params(axis='x', labelsize=12)
            ax.tick_params(axis='y', labelsize=12)
            ax.grid(axis="x", linestyle="--", alpha=0.6)
            plt.savefig(output_path, bbox_inches='tight', dpi=300)
            plt.close()
            messagebox.showinfo("完了", f"グラフを保存しました:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました:\n{str(e)}")

    def generate_wordcloud():
        try:
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("エラー", "フォントの読み込みに失敗しました。")
                return

            df = pd.read_excel(selected_excel_file, sheet_name=selected_sheet.get())
            freq_dict = pd.Series(df["count"].values, index=df["word"]).to_dict()

            wc = WordCloud(
                font_path=font_path,
                width=800,
                height=600,
                background_color="white",
                prefer_horizontal=1.0,
                colormap="tab20"
            )
            wc.generate_from_frequencies(freq_dict)

            base_path = os.path.splitext(selected_excel_file)[0]
            sheet_name = selected_sheet.get().replace(" ", "_")  # Replace spaces with underscores for file name
            output_path = f"{base_path}_{sheet_name}_wordcloud.png"
            plt.figure(figsize=(10, 8))
            plt.imshow(wc, interpolation="bilinear")
            plt.axis("off")
            plt.tight_layout(pad=0)
            plt.savefig(output_path, bbox_inches='tight', dpi=300)
            plt.close()
            messagebox.showinfo("完了", f"Word Cloudを保存しました:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました:\n{str(e)}")

    btn_graph = tk.Button(root, text="Generate Graph", command=generate_graph, width=20, height=2)
    btn_graph.pack(pady=20)

    btn_wordcloud = tk.Button(root, text="Generate Word Cloud", command=generate_wordcloud, width=20, height=2)
    btn_wordcloud.pack(pady=20)

    root.mainloop()
