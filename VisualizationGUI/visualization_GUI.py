import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import os
import pandas as pd
import matplotlib.pyplot as plt
import japanize_matplotlib
import matplotlib.font_manager as fm
from wordcloud import WordCloud
import sys

# Load bundled Japanese font (ipaexg.ttf)
def load_japanese_font():
    try:
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller EXE
            font_path = os.path.join(sys._MEIPASS, "ipaexg.ttf")
        else:
            # Running as .py script
            font_path = os.path.join(os.path.dirname(__file__), "ipaexg.ttf")
        fm.fontManager.addfont(font_path)
        font_name = fm.FontProperties(fname=font_path).get_name()
        plt.rcParams['font.family'] = font_name
        return font_path  # Return the path for WordCloud font_path param
    except Exception as e:
        print(f"Font load failed: {e}")
        return None

# Ask the user to select a CSV file before opening the UI
def select_initial_file():
    """Prompt the user to select a CSV file at the start."""
    file_path = filedialog.askopenfilename(
        title="CSVファイルを選択してください",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not file_path:
        messagebox.showinfo("キャンセル", "ファイルが選択されませんでした。プログラムを終了します。")
        sys.exit()  # Exit the program if no file is selected
    return file_path

# Main program starts here
if __name__ == "__main__":
    # Initialize Tkinter root for file dialog
    root = tk.Tk()
    root.withdraw()  # Hide the root window for the initial file selection

    # Prompt the user to select a CSV file
    selected_csv_file = select_initial_file()

    # Show the main UI after a file is selected
    root.deiconify()  # Show the root window for the main UI
    root.title("Graph and Word Cloud Generator")
    root.geometry("400x200")

    def generate_graph():
        """Generate and save the word frequency graph."""
        try:
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("エラー", "フォントの読み込みに失敗しました。")
                return
            df = pd.read_csv(selected_csv_file)

            # Sort by count
            df_sorted = df.sort_values(by="count", ascending=False)

            # Dynamically adjust figure height based on number of words
            num_words = len(df_sorted)
            height_per_word = 0.25  # adjust for vertical spacing
            base_height = 4
            fig_height = max(base_height, num_words * height_per_word)

            plt.figure(figsize=(12, fig_height))

            # Plot
            plt.barh(df_sorted["word"], df_sorted["count"], color="skyblue")
            plt.xlabel("Count", fontsize=14)
            plt.ylabel("Word", fontsize=14)
            plt.title("Word Frequency Graph", fontsize=16)
            plt.xticks(fontsize=12)
            plt.yticks(fontsize=12)
            plt.grid(axis="x", linestyle="--", alpha=0.6)
            plt.gca().invert_yaxis()

            # Save graph
            base_path = os.path.splitext(selected_csv_file)[0]
            output_path = f"{base_path}_graph.png"
            plt.savefig(output_path, bbox_inches='tight', dpi=300)
            plt.close()
            messagebox.showinfo("完了", f"グラフを保存しました:\n{output_path}")

            # Open the saved graph image
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました:\n{str(e)}")

    def generate_wordcloud():
        """Generate and save the word cloud image."""
        try:
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("エラー", "フォントの読み込みに失敗しました。")
                return

            df = pd.read_csv(selected_csv_file)
            # Convert to dictionary: {word: count}
            freq_dict = pd.Series(df["count"].values, index=df["word"]).to_dict()

            # WordCloud settings
            wc = WordCloud(
                font_path=font_path,
                width=800,
                height=600,
                background_color="white",
                prefer_horizontal=1.0,
                colormap="tab20"
            )
            wc.generate_from_frequencies(freq_dict)

            # Save word cloud
            base_path = os.path.splitext(selected_csv_file)[0]
            output_path = f"{base_path}_wordcloud.png"
            plt.figure(figsize=(10, 8))
            plt.imshow(wc, interpolation="bilinear")
            plt.axis("off")
            plt.tight_layout(pad=0)
            plt.savefig(output_path, bbox_inches='tight', dpi=300)
            plt.close()
            messagebox.showinfo("完了", f"Word Cloudを保存しました:\n{output_path}")

            # Open the saved word cloud image
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました:\n{str(e)}")

    # Create buttons
    btn_graph = tk.Button(root, text="Generate Graph", command=generate_graph, width=20, height=2)
    btn_graph.pack(pady=20)

    btn_wordcloud = tk.Button(root, text="Generate Word Cloud", command=generate_wordcloud, width=20, height=2)
    btn_wordcloud.pack(pady=20)

    # Run the application
    root.mainloop()