import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from config import (
    btn_bg, top_n_options, label_font, label_bg,
    button_font, button_padding, small_button_font, small_button_padding
)
from get_utils import get_sheet_names, get_save_path



class VisualizationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Defect Report Analysis Tool")
        self.geometry("600x500")
        # --- Styling ---
        style = ttk.Style()
        style.theme_use("xpnative")
        style.configure("TButton", font=button_font, padding=button_padding)
        style.configure("Small.TButton", font=small_button_font, padding=small_button_padding)
        # --- Variables ---
        self.selected_excel_file = None
        self.selected_sheet = tk.StringVar(self)
        self.sheet_menu = None
        self.sheet_names_global = []
        # --- Top Frame ---
        self.frame_top = tk.Frame(self, padx=10, pady=10)
        self.frame_top.pack(fill='x', anchor='nw')
        # --- Button Width ---
        button_width = 30
        # --- Add Resolution Period Column Button ---
        btn_add_resolution_period = ttk.Button(
            self.frame_top, 
            text="Add Resolution Period Column", 
            command=self.add_resolution_period_column, 
            width=button_width
        )
        btn_add_resolution_period.pack(padx=(0, 5), pady=(0, 5))
        btn_add_resolution_period.configure(style="Small.TButton")
        # --- Build Model Button ---
        btn_build_model = ttk.Button(self.frame_top, text="Build Model", command=self.build_model, width=button_width)
        btn_build_model.pack(padx=(0, 5), pady=(0, 5))
        btn_build_model.configure(style="Small.TButton")
        # --- Categorize Defects Button ---
        btn_categorize = ttk.Button(self.frame_top, text="Categorize Defects", command=self.categorize_defects, width=button_width)
        btn_categorize.pack(padx=(0, 5), pady=(0, 5))
        btn_categorize.configure(style="Small.TButton")
        # --- Generate Pie Chart Button ---
        btn_piechart = ttk.Button(self.frame_top, text="Generate Defect Category Pie Chart", command=self.generate_pie_chart, width=button_width)
        btn_piechart.pack(padx=(0, 5), pady=(0, 5))
        btn_piechart.configure(style="Small.TButton")
        # --- Generate Word Count Button ---
        btn_wordcount = ttk.Button(self.frame_top, text="Generate Word Count Table", command=self.generate_wordcount_table, width=button_width)
        btn_wordcount.pack(padx=(0, 5), pady=(0, 5))
        btn_wordcount.configure(style="Small.TButton")
        # --- Info Label ---
        self.selected_info_label = tk.Label(self.frame_top, text="", bg=label_bg, font=label_font)
        self.selected_info_label.pack(fill='x', padx=10, pady=(0, 5))
        # --- Top N Variables ---
        defaultTopNLabel = 'Top 10'
        defaultTopN = 10
        self.selected_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_top_n = tk.IntVar(value=defaultTopN)
        self.selected_wc_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_wc_top_n = tk.IntVar(value=defaultTopN)
        self.selected_bubble_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_bubble_top_n = tk.IntVar(value=defaultTopN)
        # --- Buttons and Menus ---
        self.setup_gui()

    def add_resolution_period_column(self):
        file_path = filedialog.askopenfilename(
            title="Select a Defect Report", 
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "No file selected.")
            return
        try:
            from resolution_column import add_resolution_period_column_logic
            add_resolution_period_column_logic(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add resolution period column:\n{str(e)}")
    

    def build_model(self):
        file_path = filedialog.askopenfilename(
            title="Select Train Data",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "No file selected for model training.")
            return
        try:
            from build_model import train_model
            train_model(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to build model:\n{str(e)}")

    def categorize_defects(self):
        file_path = filedialog.askopenfilename(
            title="Select Defect Report to Categorize",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "No defect report selected for categorization.")
            return
        try:
            from summary_classifier import categorize_summaries
            categorize_summaries(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to categorize defects:\n{str(e)}")

    def generate_pie_chart(self):
        file_path = filedialog.askopenfilename(
            title="Select Categorized Defect Report for Pie Chart",
            filetypes=[("Excel files", "*categorized.xlsx;*categorized.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "No categorized defect report selected for pie chart.")
            return
        try:
            from plots import generate_category_pie_chart
            generate_category_pie_chart(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate pie chart:\n{str(e)}")

    def generate_wordcount_table(self):
        file_path = filedialog.askopenfilename(
            title="Select Defect Report for Word Count",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "No defect report selected for word count.")
            return
        try:
            from word_count_util import word_count
            word_count(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate a word count table:\n{str(e)}")


    def setup_gui(self):
        # --- Select File Button ---
        btn_select_file = ttk.Button(self.frame_top, text="Select Word Count Table", command=self.select_excel_file)
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
        top_n_menu.config(font=button_font, bg=btn_bg, highlightthickness=1, relief="groove")
        top_n_menu["menu"].config(font=button_font, bg=btn_bg)
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
        wc_top_n_menu.config(font=button_font, bg=btn_bg, highlightthickness=1, relief="groove")
        wc_top_n_menu["menu"].config(font=button_font, bg=btn_bg)
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
        bubble_top_n_menu.config(font=button_font, bg=btn_bg, highlightthickness=1, relief="groove")
        bubble_top_n_menu["menu"].config(font=button_font, bg=btn_bg)
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
            title="Select a Word Count Table",
            filetypes=[("Excel files", "*wordcount*.xlsx;*wordcount*.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "There is no file selected")
            return
        self.selected_excel_file = file_path
        try:
            sheet_names = get_sheet_names(self.selected_excel_file)
            self.sheet_names_global = sheet_names
            self._missing_col_error_shown_per_sheet = {}  # Track error per sheet
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
                self.sheet_menu.config(font=small_button_font, bg=btn_bg, highlightthickness=1, relief="groove")
                self.sheet_menu["menu"].config(font=small_button_font, bg=btn_bg)
                self.sheet_menu.pack(side='right', anchor='ne', padx=10, pady=(0, 10))
                self.selected_info_label.config(
                    text=f"File: {os.path.basename(self.selected_excel_file)} | Category: {self.selected_sheet.get()}"
                )
                def update_label(*args):
                    self.selected_info_label.config(
                        text=f"File: {os.path.basename(self.selected_excel_file)} | Category: {self.selected_sheet.get()}"
                    )
                    try:
                        import pandas as pd
                        self.df = pd.read_excel(self.selected_excel_file, sheet_name=self.selected_sheet.get())
                        sheet = self.selected_sheet.get()
                        # Normalize column names to lowercase
                        self.df.columns = [col.lower() for col in self.df.columns]
                        if "word" not in self.df.columns or "count" not in self.df.columns:
                            if not self._missing_col_error_shown_per_sheet.get(sheet, False):
                                messagebox.showerror("Missing Columns", "The selected file does not contain both 'Word' and 'Count' columns.")
                                self._missing_col_error_shown_per_sheet[sheet] = True
                            self.selected_info_label.config(text="")
                            if self.sheet_menu:
                                self.sheet_menu.destroy()
                                self.sheet_menu = None
                            self.df = None
                        else:
                            self._missing_col_error_shown_per_sheet[sheet] = False
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to read sheet:\n{str(e)}")
                        self.df = None
                self.selected_sheet.trace_add('write', update_label)
            # Create dataframe for the first sheet by default
            import pandas as pd
            self.df = pd.read_excel(self.selected_excel_file, sheet_name=self.selected_sheet.get())
            sheet = self.selected_sheet.get()
            # Normalize column names to lowercase
            self.df.columns = [col.lower() for col in self.df.columns]
            if "word" not in self.df.columns or "count" not in self.df.columns:
                if not self._missing_col_error_shown_per_sheet.get(sheet, False):
                    messagebox.showerror("Error", "There is no 'Word' or 'Count' column in the data")
                    self._missing_col_error_shown_per_sheet[sheet] = True
                self.selected_info_label.config(text="")
                if self.sheet_menu:
                    self.sheet_menu.destroy()
                    self.sheet_menu = None
                self.df = None
            else:
                self._missing_col_error_shown_per_sheet[sheet] = False
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets:\n{str(e)}")

    def on_generate_graph(self):
        if not self.selected_excel_file or self.df is None:
            messagebox.showwarning("No File Selected", "Please select a word count table first.")
            return
        try:
            from plots import load_japanese_font, get_colordict, generate_graph
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("Error", "Failed to read fonts")
                return
            n = self.selected_top_n.get()
            actual_n = min(n, len(self.df))
            if len(self.df) < n:
                messagebox.showwarning(
                    "Not Enough Words",
                    f"The selected word count table does not have enough words. Showing {actual_n} words instead of {n}."
                )
            df_top = self.df.iloc[:actual_n]
            chunk_size = 20
            max_count = df_top['count'].max()
            color_dict = get_colordict('viridis', max_count, 1)
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            top_n_label = f"Top{actual_n}" if actual_n != n else self.selected_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{top_n_label}_graph")
            generate_graph(df_top, chunk_size, color_dict, f"{top_n_label}", sheet_name, output_path, self.sheet_names_global)
            messagebox.showinfo("Completed", f"Saved a graph:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("Error", f"There was an error:\n{str(e)}")

    def on_generate_wordcloud(self):
        if not self.selected_excel_file or self.df is None:
            messagebox.showwarning("No File Selected", "Please select a word count table first.")
            return
        try:
            from plots import load_japanese_font, generate_wordcloud
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("Error", "Failed to read fonts")
                return
            n = self.selected_wc_top_n.get()
            actual_n = min(n, len(self.df))
            if len(self.df) < n:
                messagebox.showwarning(
                    "Not Enough Words",
                    f"The selected word count table does not have enough words. Showing {actual_n} words instead of {n}."
                )
            wc_top_n_label = f"Top{actual_n}" if actual_n != n else self.selected_wc_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{wc_top_n_label}_wordcloud")
            generate_wordcloud(self.df.iloc[:actual_n], actual_n, font_path, output_path, sheet_name, wc_top_n_label, self.sheet_names_global)
            messagebox.showinfo("Completed", f"Saved a word cloud:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("Error", f"There was an error:\n{str(e)}")

    def on_generate_bubble_chart(self):
        if not self.selected_excel_file or self.df is None:
            messagebox.showwarning("No File Selected", "Please select a word count table first.")
            return
        try:
            from plots import load_japanese_font, get_colordict, generate_bubble_chart
            font_path = load_japanese_font()
            if not font_path:
                messagebox.showerror("Error", "Failed to read fonts")
                return
            n = self.selected_bubble_top_n.get()
            actual_n = min(n, len(self.df))
            if len(self.df) < n:
                messagebox.showwarning(
                    "Not Enough Words",
                    f"The selected word count table does not have enough words. Showing {actual_n} words instead of {n}."
                )
            bubble_top_n_label = f"Top{actual_n}" if actual_n != n else self.selected_bubble_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            max_count = self.df['count'].max()
            color_dict = get_colordict('summer', max_count, min(self.df['count']))
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{bubble_top_n_label}_bubblechart")
            generate_bubble_chart(self.df.iloc[:actual_n], actual_n, color_dict, output_path, sheet_name, bubble_top_n_label,self.sheet_names_global)
            messagebox.showinfo("Completed", f"Saved a bubble chart:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("Error", f"There was an error:\n{str(e)}")