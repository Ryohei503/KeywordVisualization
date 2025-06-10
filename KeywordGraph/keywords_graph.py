import pandas as pd
import matplotlib.pyplot as plt
import japanize_matplotlib
import matplotlib.font_manager as fm
from tkinter import filedialog, Tk, messagebox
import os
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
    except Exception as e:
        print(f"Font load failed: {e}")

def select_file():
    """Open a file dialog to select a CSV file"""
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="CSVファイルを選択してください",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    return file_path

def generate_graph(csv_path):
    """Generate and save the word frequency graph"""
    df = pd.read_csv(csv_path)

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
    base_path = os.path.splitext(csv_path)[0]
    output_path = f"{base_path}_graph.png"
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    return output_path

def main():
    load_japanese_font()

    print("CSVファイルを選択してください...")

    csv_path = select_file()

    if not csv_path:
        messagebox.showinfo("キャンセル", "ファイルが選択されませんでした。終了します。")
        return

    try:
        output_path = generate_graph(csv_path)
        messagebox.showinfo("完了", f"グラフを保存しました:\n{output_path}")
    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました:\n{str(e)}")

if __name__ == "__main__":
    main()
