import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import os
import pandas as pd
from wordcloud import WordCloud
import sys
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.colors import Normalize
import numpy as np
import seaborn as sns

# Apply Seaborn darkgrid style globally
sns.set_style("darkgrid")

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


# Get sheet names from the selected Excel file
def get_sheet_names(file_path):
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names

# Generate color dictionary
def get_colordict(cmap_name, max_value, min_value=1):
    cmap = plt.get_cmap(cmap_name)
    norm = Normalize(vmin=min_value, vmax=max_value)
    return {i: cmap(norm(i)) for i in range(min_value, max_value + 1)}

# GUI setup
root = tk.Tk()
root.title("Defect Summary Analysis Tool")
root.geometry("400x400")

selected_excel_file = None
selected_sheet = tk.StringVar(root)
sheet_menu = None

def select_excel_file():
    global selected_excel_file, sheet_menu
    file_path = filedialog.askopenfilename(
        title="Select an Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
    )
    if not file_path:
        messagebox.showinfo("Cancel", "There is no file selected")
        return
    selected_excel_file = file_path
    try:
        sheet_names = get_sheet_names(selected_excel_file)
        selected_sheet.set(sheet_names[0])
        if sheet_menu:
            sheet_menu.destroy()
        sheet_menu = tk.OptionMenu(root, selected_sheet, *sheet_names)
        sheet_menu.pack(anchor='nw', padx=10, pady=(50, 10))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read sheets:\n{str(e)}")


def generate_graph():
    try:
        font_path = load_japanese_font()
        if not font_path:
            messagebox.showerror("Error", "Failed to read fonts")
            return

        # Load and sort data
        df = pd.read_excel(selected_excel_file, sheet_name=selected_sheet.get())
        # Ensure the required columns exist
        if "word" not in df.columns or "count" not in df.columns:
            messagebox.showerror("Error", "There is no 'word' or 'count' column in the data")
            return
        df_sorted = df.sort_values(by="count", ascending=False).reset_index(drop=True)

        # Create index list for slicing  
        num_words = len(df_sorted)
        num_chunks = max(1, (num_words + 19) // 20) # One chunk per 20 words, minimum 1

        index_list = [[i[0], i[-1] + 1] for i in np.array_split(range(len(df_sorted)), num_chunks)]

        # Get color dictionary
        max_count = df_sorted['count'].max()
        color_dict = get_colordict('viridis', max_count, 1)

        # Create subplots
        fig, axs = plt.subplots(1, num_chunks, figsize=(16, 8), facecolor='white', squeeze=False)

        for col, idx in zip(range(num_chunks), index_list):
            df_chunk = df_sorted.iloc[idx[0]:idx[1]]
            labels = [f"{w}: {n}" for w, n in zip(df_chunk['word'], df_chunk['count'])]
            colors = [color_dict.get(i) for i in df_chunk['count']]
            x = list(df_chunk['count'])
            y = list(range(len(df_chunk)))

            sns.barplot(
                x=x, y=y, data=df_chunk, alpha=0.9, orient='h',
                ax=axs[0][col], palette=colors
            )
            axs[0][col].set_xlim(0, max_count + 1)  # Set X-axis range
            axs[0][col].set_yticklabels(labels, fontsize=12)
            axs[0][col].spines['bottom'].set_color('white')
            axs[0][col].spines['right'].set_color('white')
            axs[0][col].spines['top'].set_color('white')
            axs[0][col].spines['left'].set_color('white')
            # Move the x-axis to top
            ax=axs[0][col]
            ax.xaxis.set_ticks_position('top')
            ax.xaxis.set_label_position('top')


        # Add a title
        sheet_name = selected_sheet.get().replace(" ", "_")
        fig.    suptitle(f"Word Frequency Graph - {sheet_name}", fontsize=16)
        plt.tight_layout()
        

        # Save and display the graph
        base_path = os.path.splitext(selected_excel_file)[0]
        output_path = f"{base_path}_{sheet_name}_graph.png"
        plt.savefig(output_path, bbox_inches='tight', dpi=300)
        plt.close()

        messagebox.showinfo("Completed", f"Saved a graph:\n{output_path}")
        os.startfile(output_path)

    except Exception as e:
        messagebox.showerror("Error", f"There was an error:\n{str(e)}")

def generate_wordcloud():
    try:
        font_path = load_japanese_font()
        if not font_path:
            messagebox.showerror("Error", "Failed to read fonts")
            return

        df = pd.read_excel(selected_excel_file, sheet_name=selected_sheet.get())
        # Ensure the required columns exist
        if "word" not in df.columns or "count" not in df.columns:
            messagebox.showerror("Error", "There is no 'word' or 'count' column in the data")
            return
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
        messagebox.showinfo("Completed", f"Saved a word cloud:\n{output_path}")
        os.startfile(output_path)
    except Exception as e:
        messagebox.showerror("Error", f"There was an error:\n{str(e)}")


def generate_bubble_chart():
    try:
        import circlify
        font_path = load_japanese_font()
        if not font_path:
            messagebox.showerror("Error", "Failed to read fonts")
            return

        # Load the data from the selected Excel file
        df = pd.read_excel(selected_excel_file, sheet_name=selected_sheet.get())

        # Ensure the required columns exist
        if "word" not in df.columns or "count" not in df.columns:
            messagebox.showerror("Error", "There is no 'word' or 'count' column in the data")
            return

        # Prepare data
        df_words = df.sort_values(by="count", ascending=False).reset_index(drop=True)
        labels = list(df_words['word'])
        counts = list(df_words['count'])

        # Reverse for display order (largest at bottom)
        labels.reverse()
        counts.reverse()

        # Generate circles
        circles = circlify.circlify(
            counts,
            show_enclosure=False,
            target_enclosure=circlify.Circle(x=0, y=0, r=1)
        )

        # Get color dictionary
        max_count = max(counts)
        color_dict = get_colordict('summer', max_count, min(counts))

        # Plot
        fig, ax = plt.subplots(figsize=(9, 9), facecolor='white')
        ax.axis('off')
        lim = max(
            max(abs(circle.x) + circle.r, abs(circle.y) + circle.r)
            for circle in circles
        )
        plt.xlim(-lim, lim)
        plt.ylim(-lim, lim)

        # Draw circles and annotate
        for circle, label, count in zip(circles, labels, counts):
            x, y, r = circle.x, circle.y, circle.r
            ax.add_patch(plt.Circle((x, y), r, alpha=0.9, color=color_dict.get(count)))
            plt.annotate(f"{label}\n{count}", (x, y), size=12, va='center', ha='center')

        # Save and display the bubble chart
        base_path = os.path.splitext(selected_excel_file)[0]
        sheet_name = selected_sheet.get().replace(" ", "_")  # Replace spaces with underscores for file name
        output_path = f"{base_path}_{sheet_name}_bubblechart.png"
        plt.title("Keyword Frequency Bubble Chart")
        plt.tight_layout()
        plt.savefig(output_path, bbox_inches='tight', dpi=300)
        plt.close()
        messagebox.showinfo("Completed", f"Saved a bubble chart:\n{output_path}")
        os.startfile(output_path)
        
    except Exception as e:
        messagebox.showerror("Error", f"There was an error when generating a bubble chart:\n{str(e)}")


# GUI Buttons
btn_select_file = tk.Button(root, text="Select Excel File", command=select_excel_file, width=20, height=1)
btn_select_file.pack(anchor='nw')

btn_graph = tk.Button(root, text="Generate Graph", command=generate_graph, width=20, height=2)
btn_graph.pack(pady=20)

btn_wordcloud = tk.Button(root, text="Generate Word Cloud", command=generate_wordcloud, width=20, height=2)
btn_wordcloud.pack(pady=20)

btn_bubble = tk.Button(root, text="Generate Bubble Chart", command=generate_bubble_chart, width=20, height=2)
btn_bubble.pack(pady=20)

root.mainloop()

