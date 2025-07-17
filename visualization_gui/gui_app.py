import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import pandas as pd
from config import top_n_menu_font, sheet_menu_font, optionmenu_bg, top_n_options
from get_utils import get_sheet_names, get_save_path
from plots import (
    load_japanese_font, get_colordict,
    generate_graph, generate_wordcloud, generate_bubble_chart
)

class VisualizationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Defect Report Analysis Tool")
        self.geometry("600x400")
        # --- Styling ---
        style = ttk.Style()
        style.theme_use("xpnative")
        style.configure("TButton", font=("Arial", 12), padding=5)
        style.configure("Small.TButton", font=("Arial", 10), padding=5)
        # --- Variables ---
        self.selected_excel_file = None
        self.selected_sheet = tk.StringVar(self)
        self.sheet_menu = None
        self.sheet_names_global = []
        # --- Top Frame ---
        self.frame_top = tk.Frame(self, padx=10, pady=10)
        self.frame_top.pack(fill='x', anchor='nw')
        self.selected_info_label = tk.Label(self.frame_top, text="", bg="white", font=("Arial", 10), anchor='w')
        self.selected_info_label.pack(fill='x', padx=10, pady=(0, 5))
        # --- Top N Variables ---
        self.selected_top_n_label = tk.StringVar(value="Top 20")
        self.selected_top_n = tk.IntVar(value=20)
        self.selected_wc_top_n_label = tk.StringVar(value="Top 20")
        self.selected_wc_top_n = tk.IntVar(value=20)
        self.selected_bubble_top_n_label = tk.StringVar(value="Top 20")
        self.selected_bubble_top_n = tk.IntVar(value=20)
        # --- Buttons and Menus ---
        self.setup_gui()

    def setup_gui(self):
        # --- Select File Button ---
        btn_select_file = ttk.Button(self.frame_top, text="Select Excel File", command=self.select_excel_file)
        btn_select_file.pack(side='left', padx=(0, 10), pady=5)
        btn_select_file.configure(style="Small.TButton")
        # --- Frame for Graph/WordCloud/Bubble ---
        frame_buttons = tk.Frame(self, padx=20, pady=20)
        frame_buttons.pack(fill='both', expand=True)
        # --- Graph ---
        frame_graph = tk.Frame(frame_buttons)
        frame_graph.pack(pady=10)
        btn_graph = ttk.Button(frame_graph, text="Generate Graph", command=self.on_generate_graph, width=20)
        btn_graph.pack(side='left', padx=(0, 10))
        top_n_menu = tk.OptionMenu(
            frame_graph,
            self.selected_top_n_label,
            *[label for label, val in top_n_options],
            command=lambda label: self.set_top_n(dict(top_n_options)[label], label)
        )
        top_n_menu.config(font=top_n_menu_font, bg=optionmenu_bg, highlightthickness=1, relief="groove")
        top_n_menu["menu"].config(font=top_n_menu_font, bg=optionmenu_bg)
        top_n_menu.pack(side='left')
        # --- Word Cloud ---
        frame_wordcloud = tk.Frame(frame_buttons)
        frame_wordcloud.pack(pady=10)
        btn_wordcloud = ttk.Button(frame_wordcloud, text="Generate Word Cloud", command=self.on_generate_wordcloud, width=20)
        btn_wordcloud.pack(side='left', padx=(0, 10))
        wc_top_n_menu = tk.OptionMenu(
            frame_wordcloud,
            self.selected_wc_top_n_label,
            *[label for label, val in top_n_options],
            command=lambda label: self.set_wc_top_n(dict(top_n_options)[label], label)
        )
        wc_top_n_menu.config(font=top_n_menu_font, bg=optionmenu_bg, highlightthickness=1, relief="groove")
        wc_top_n_menu["menu"].config(font=top_n_menu_font, bg=optionmenu_bg)
        wc_top_n_menu.pack(side='left')
        # --- Bubble Chart ---
        frame_bubble = tk.Frame(frame_buttons)
        frame_bubble.pack(pady=10)
        btn_bubble = ttk.Button(frame_bubble, text="Generate Bubble Chart", command=self.on_generate_bubble_chart, width=20)
        btn_bubble.pack(side='left', padx=(0, 10))
        bubble_top_n_menu = tk.OptionMenu(
            frame_bubble,
            self.selected_bubble_top_n_label,
            *[label for label, val in top_n_options],
            command=lambda label: self.set_bubble_top_n(dict(top_n_options)[label], label)
        )
        bubble_top_n_menu.config(font=top_n_menu_font, bg=optionmenu_bg, highlightthickness=1, relief="groove")
        bubble_top_n_menu["menu"].config(font=top_n_menu_font, bg=optionmenu_bg)
        bubble_top_n_menu.pack(side='left')

    def set_top_n(self, val, label):
        self.selected_top_n.set(val)
        self.selected_top_n_label.set(label)

    def set_wc_top_n(self, val, label):
        self.selected_wc_top_n.set(val)
        self.selected_wc_top_n_label.set(label)

    def set_bubble_top_n(self, val, label):
        self.selected_bubble_top_n.set(val)
        self.selected_bubble_top_n_label.set(label)

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Select an Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if not file_path:
            messagebox.showinfo("Cancel", "There is no file selected")
            return
        self.selected_excel_file = file_path
        try:
            sheet_names = get_sheet_names(self.selected_excel_file)
            self.sheet_names_global = sheet_names
            # Remove previous OptionMenu if it exists
            if self.sheet_menu:
                self.sheet_menu.destroy()
                self.sheet_menu = None
            if len(sheet_names) == 1:
                self.selected_sheet.set(sheet_names[0])
                self.selected_info_label.config(
                    text=f"File: {os.path.basename(self.selected_excel_file)} | Category: None"
                )
            else:
                self.selected_sheet.set(sheet_names[0])
                self.sheet_menu = tk.OptionMenu(self.frame_top, self.selected_sheet, *sheet_names)
                self.sheet_menu.config(font=sheet_menu_font, bg=optionmenu_bg, highlightthickness=1, relief="groove")
                self.sheet_menu["menu"].config(font=sheet_menu_font, bg=optionmenu_bg)
                self.sheet_menu.pack(side='right', anchor='ne', padx=10, pady=(0, 10))
                self.selected_info_label.config(
                    text=f"File: {os.path.basename(self.selected_excel_file)} | Category: {self.selected_sheet.get()}"
                )
                def update_label(*args):
                    self.selected_info_label.config(
                        text=f"File: {os.path.basename(self.selected_excel_file)} | Category: {self.selected_sheet.get()}"
                    )
                self.selected_sheet.trace_add('write', update_label)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets:\n{str(e)}")

    def on_generate_graph(self):
        if not self.selected_excel_file:
            messagebox.showwarning("No File Selected", "Please select an Excel file first.")
            return
        try:
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("Error", "Failed to read fonts")
                return
            df = pd.read_excel(self.selected_excel_file, sheet_name=self.selected_sheet.get())
            if "word" not in df.columns or "count" not in df.columns:
                messagebox.showerror("Error", "There is no 'word' or 'count' column in the data")
                return
            df_sorted = df.sort_values(by="count", ascending=False).reset_index(drop=True)
            n = self.selected_top_n.get()
            df_sorted = df_sorted.iloc[:n]
            chunk_size = 20
            max_count = df_sorted['count'].max()
            color_dict = get_colordict('viridis', max_count, 1)
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            file_name = os.path.basename(self.selected_excel_file)
            top_n_label = self.selected_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{top_n_label}_graph")
            generate_graph(df_sorted, n, chunk_size, color_dict, top_n_label, sheet_name, file_name, output_path, self.sheet_names_global)
            messagebox.showinfo("Completed", f"Saved a graph:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("Error", f"There was an error:\n{str(e)}")

    def on_generate_wordcloud(self):
        if not self.selected_excel_file:
            messagebox.showwarning("No File Selected", "Please select an Excel file first.")
            return
        try:
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("Error", "Failed to read fonts")
                return
            df = pd.read_excel(self.selected_excel_file, sheet_name=self.selected_sheet.get())
            if "word" not in df.columns or "count" not in df.columns:
                messagebox.showerror("Error", "There is no 'word' or 'count' column in the data")
                return
            n = self.selected_wc_top_n.get()
            wc_top_n_label = self.selected_wc_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{wc_top_n_label}_wordcloud")
            generate_wordcloud(df, n, font_path, wc_top_n_label, output_path)
            messagebox.showinfo("Completed", f"Saved a word cloud:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("Error", f"There was an error:\n{str(e)}")

    def on_generate_bubble_chart(self):
        if not self.selected_excel_file:
            messagebox.showwarning("No File Selected", "Please select an Excel file first.")
            return
        try:
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("Error", "Failed to read fonts")
                return
            df = pd.read_excel(self.selected_excel_file, sheet_name=self.selected_sheet.get())
            if "word" not in df.columns or "count" not in df.columns:
                messagebox.showerror("Error", "There is no 'word' or 'count' column in the data")
                return
            n = self.selected_bubble_top_n.get()
            bubble_top_n_label = self.selected_bubble_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            max_count = df['count'].max()
            color_dict = get_colordict('summer', max_count, min(df['count']))
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{bubble_top_n_label}_bubblechart")
            generate_bubble_chart(df, n, color_dict, bubble_top_n_label, output_path)
            messagebox.showinfo("Completed", f"Saved a bubble chart:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("Error", f"There was an error when generating a bubble chart:\n{str(e)}")