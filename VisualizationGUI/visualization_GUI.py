import tkinter as tk
from tkinter import ttk
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
root.title("Defect Report Analysis Tool")
root.geometry("600x400")
frame_top = tk.Frame(root, padx=10, pady=10)
frame_top.pack(fill='x', anchor='nw')
frame_buttons = tk.Frame(root, padx=20, pady=20)
frame_buttons.pack(fill='both', expand=True)

selected_excel_file = None
selected_sheet = tk.StringVar(root)
sheet_menu = None
sheet_names_global = []

# Add this after defining frame_top and before defining select_excel_file
selected_info_label = tk.Label(frame_top, text="", bg="white", font=("Arial", 10), anchor='w')
selected_info_label.pack(fill='x', padx=10, pady=(0, 5))

# Add this global variable and dropdown options near the top of your file
top_n_options = [("Top 20", 20), ("Top 40", 40), ("Top 60", 60), ("Top 80", 80), ("Top 100", 100)]
selected_top_n_label = tk.StringVar(value="Top 20")
selected_top_n = tk.IntVar(value=20)

def set_top_n(val, label):
    selected_top_n.set(val)
    selected_top_n_label.set(label)

def select_excel_file():
    global selected_excel_file, sheet_menu, sheet_names_global
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
        sheet_names_global = sheet_names  # Store globally for later use
        # Remove previous OptionMenu if it exists
        if sheet_menu:
            sheet_menu.destroy()
            sheet_menu = None
        if len(sheet_names) == 1:
            selected_sheet.set(sheet_names[0])
            # Show 'Category: None' if only one sheet
            selected_info_label.config(
                text=f"File: {os.path.basename(selected_excel_file)} | Category: None"
            )
        else:
            selected_sheet.set(sheet_names[0])
            sheet_menu = tk.OptionMenu(frame_top, selected_sheet, *sheet_names)
            sheet_menu.pack(anchor='ne', padx=10, pady=(0, 10))
            # Show first sheet as category
            selected_info_label.config(
                text=f"File: {os.path.basename(selected_excel_file)} | Category: {selected_sheet.get()}"
            )
            # Update label when the sheet selection changes
            def update_label(*args):
                selected_info_label.config(
                    text=f"File: {os.path.basename(selected_excel_file)} | Category: {selected_sheet.get()}"
                )
            selected_sheet.trace_add('write', update_label)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read sheets:\n{str(e)}")

def get_save_path(base_path, sheet_names, sheet_name, suffix):
    # If only one sheet, do not include sheet name in file name
    if len(sheet_names) == 1:
        return f"{base_path}_{suffix}.png"
    else:
        return f"{base_path}_{sheet_name}_{suffix}.png"

def generate_graph():
    if not selected_excel_file:
        messagebox.showwarning("No File Selected", "Please select an Excel file first.")
        return
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

        # Only use the top N words as selected by the dropdown
        n = selected_top_n.get()
        df_sorted = df_sorted.iloc[:n]

        # Create index list for slicing  
        num_words = len(df_sorted)
        chunk_size = 20
        index_list = [
            [start, min(start + chunk_size, num_words)]
            for start in range(0, num_words, chunk_size)
        ]
        num_chunks = len(index_list)

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

        # Add a title with a category and selected top_n_menu option
        sheet_name = selected_sheet.get().replace(" ", "_")
        file_name = os.path.basename(selected_excel_file)
        top_n_label = selected_top_n_label.get().replace(" ", "")  # e.g., "Top20"
        # Only include sheet name in title if there are multiple sheets
        if len(sheet_names_global) == 1:
            fig.suptitle(
                f"Word Frequency Graph - {top_n_label}",
                fontsize=16
            )
        else:
            fig.suptitle(
                f"Word Frequency Graph - {sheet_name} - {top_n_label}",
                fontsize=16
            )
        plt.tight_layout()

        # Save and display the graph
        base_path = os.path.splitext(selected_excel_file)[0]
        output_path = get_save_path(base_path, sheet_names_global, sheet_name, f"{top_n_label}_graph")
        plt.savefig(output_path, bbox_inches='tight', dpi=300)
        plt.close()

        messagebox.showinfo("Completed", f"Saved a graph:\n{output_path}")
        os.startfile(output_path)

    except Exception as e:
        messagebox.showerror("Error", f"There was an error:\n{str(e)}")

# --- Word Cloud Top N ---
selected_wc_top_n_label = tk.StringVar(value="Top 20")
selected_wc_top_n = tk.IntVar(value=20)

def set_wc_top_n(val, label):
    selected_wc_top_n.set(val)
    selected_wc_top_n_label.set(label)

def generate_wordcloud():
    if not selected_excel_file:
        messagebox.showwarning("No File Selected", "Please select an Excel file first.")
        return
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

        # Only use the top N words as selected by the dropdown for word cloud
        n = selected_wc_top_n.get()
        df = df.sort_values(by="count", ascending=False).reset_index(drop=True)
        df = df.iloc[:n]

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
        wc_top_n_label = selected_wc_top_n_label.get().replace(" ", "")
        output_path = get_save_path(base_path, sheet_names_global, sheet_name, f"{wc_top_n_label}_wordcloud")
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

# --- Bubble Chart Top N ---
selected_bubble_top_n_label = tk.StringVar(value="Top 20")
selected_bubble_top_n = tk.IntVar(value=20)

def set_bubble_top_n(val, label):
    selected_bubble_top_n.set(val)
    selected_bubble_top_n_label.set(label)

def generate_bubble_chart():
    if not selected_excel_file:
        messagebox.showwarning("No File Selected", "Please select an Excel file first.")
        return
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

        # Only use the top N words as selected by the dropdown for bubble chart
        n = selected_bubble_top_n.get()
        df_words = df.sort_values(by="count", ascending=False).reset_index(drop=True)
        df_words = df_words.iloc[:n]
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
        bubble_top_n_label = selected_bubble_top_n_label.get().replace(" ", "")
        output_path = get_save_path(base_path, sheet_names_global, sheet_name, f"{bubble_top_n_label}_bubblechart")
        plt.title("Keyword Frequency Bubble Chart")
        plt.tight_layout()
        plt.savefig(output_path, bbox_inches='tight', dpi=300)
        plt.close()
        messagebox.showinfo("Completed", f"Saved a bubble chart:\n{output_path}")
        os.startfile(output_path)
    except Exception as e:
        messagebox.showerror("Error", f"There was an error when generating a bubble chart:\n{str(e)}")

# Styling
style = ttk.Style()
style.theme_use("xpnative")
style.configure("TButton", font=("Arial", 12), padding=5)
style.configure("Small.TButton", font=("Arial", 10), padding=5)  # Define the smaller style

# Style for OptionMenu (top_n_menu)
optionmenu_font = ("Arial", 12)
optionmenu_bg = "white"

# GUI Buttons
btn_select_file = ttk.Button(frame_top, text="Select Excel File", command=select_excel_file)
btn_select_file.pack(side='left', padx=(0, 10), pady=5)
btn_select_file.configure(style="Small.TButton")

# --- Graph Frame ---
frame_graph = tk.Frame(frame_buttons)
frame_graph.pack(pady=10)

btn_graph = ttk.Button(frame_graph, text="Generate Graph", command=generate_graph, width=20)
btn_graph.pack(side='left', padx=(0, 10))

top_n_menu = tk.OptionMenu(
    frame_graph,
    selected_top_n_label,
    *[label for label, val in top_n_options],
    command=lambda label: set_top_n(dict(top_n_options)[label], label)
)
top_n_menu.config(font=optionmenu_font, bg=optionmenu_bg, highlightthickness=1, relief="groove")
top_n_menu["menu"].config(font=optionmenu_font, bg=optionmenu_bg)
top_n_menu.pack(side='left')

# --- Word Cloud Frame ---
frame_wordcloud = tk.Frame(frame_buttons)
frame_wordcloud.pack(pady=10)

btn_wordcloud = ttk.Button(frame_wordcloud, text="Generate Word Cloud", command=generate_wordcloud, width=20)
btn_wordcloud.pack(side='left', padx=(0, 10))

wc_top_n_menu = tk.OptionMenu(
    frame_wordcloud,
    selected_wc_top_n_label,
    *[label for label, val in top_n_options],
    command=lambda label: set_wc_top_n(dict(top_n_options)[label], label)
)
wc_top_n_menu.config(font=optionmenu_font, bg=optionmenu_bg, highlightthickness=1, relief="groove")
wc_top_n_menu["menu"].config(font=optionmenu_font, bg=optionmenu_bg)
wc_top_n_menu.pack(side='left')

# --- Bubble Chart Frame ---
frame_bubble = tk.Frame(frame_buttons)
frame_bubble.pack(pady=10)

btn_bubble = ttk.Button(frame_bubble, text="Generate Bubble Chart", command=generate_bubble_chart, width=20)
btn_bubble.pack(side='left', padx=(0, 10))

bubble_top_n_menu = tk.OptionMenu(
    frame_bubble,
    selected_bubble_top_n_label,
    *[label for label, val in top_n_options],
    command=lambda label: set_bubble_top_n(dict(top_n_options)[label], label)
)
bubble_top_n_menu.config(font=optionmenu_font, bg=optionmenu_bg, highlightthickness=1, relief="groove")
bubble_top_n_menu["menu"].config(font=optionmenu_font, bg=optionmenu_bg)
bubble_top_n_menu.pack(side='left')

root.mainloop()

