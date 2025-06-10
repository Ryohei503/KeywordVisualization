import pandas as pd
import matplotlib.pyplot as plt
import japanize_matplotlib
import matplotlib.font_manager as fm
from tkinter import filedialog, Tk, messagebox
from wordcloud import WordCloud
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
        return font_path  # Return the path for WordCloud font_path param
    except Exception as e:
        print(f"Font load failed: {e}")
        return None

def select_file():
    """Open a file dialog to select a CSV file"""
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="CSVファイルを選択してください",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    return file_path

def generate_wordcloud(csv_path, font_path):
    """Generate and save the word cloud image"""
    df = pd.read_csv(csv_path)
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
    base_path = os.path.splitext(csv_path)[0]
    output_path = f"{base_path}_wordcloud.png"
    plt.figure(figsize=(10, 8))
    plt.imshow(wc, interpolation="bilinear")
    plt.axis("off")
    plt.tight_layout(pad=0)
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    return output_path

def main():
    font_path = load_japanese_font()

    print("CSVファイルを選択してください...")

    csv_path = select_file()

    if not csv_path:
        messagebox.showinfo("キャンセル", "ファイルが選択されませんでした。終了します。")
        return

    try:
        output_path = generate_wordcloud(csv_path, font_path)
        messagebox.showinfo("完了", f"ワードクラウド画像を保存しました:\n{output_path}")
    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました:\n{str(e)}")

if __name__ == "__main__":
    main()