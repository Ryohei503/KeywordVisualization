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
        self.geometry("600x700")
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
        button_width = 40
        # --- Filter Defect Reports Button ---
        btn_filter_defects = ttk.Button(
            self.frame_top,
            text="Filter Defect Reports",
            command=self.filter_defect_reports,
            width=button_width
        )
        btn_filter_defects.pack(padx=(0, 5), pady=(0, 5))
        btn_filter_defects.configure(style="Small.TButton")
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
        btn_build_model = ttk.Button(self.frame_top, text="Build Model for Categorization", command=self.build_model, width=button_width)
        btn_build_model.pack(padx=(0, 5), pady=(0, 5))
        btn_build_model.configure(style="Small.TButton")
        # --- Categorize Defects Button ---
        btn_categorize = ttk.Button(self.frame_top, text="Categorize Defect Reports", command=self.categorize_defects, width=button_width)
        btn_categorize.pack(padx=(0, 5), pady=(0, 5))
        btn_categorize.configure(style="Small.TButton")
        # --- Generate Pie Chart Button ---
        btn_piechart = ttk.Button(self.frame_top, text="Generate Defect Category Pie Chart", command=self.generate_pie_chart, width=button_width)
        btn_piechart.pack(padx=(0, 5), pady=(0, 5))
        btn_piechart.configure(style="Small.TButton")
        # --- Generate Box Plot Button (Priority selection via popup) ---
        btn_generate_box_plot = ttk.Button(self.frame_top, text="Generate Defect Category Box Plot", command=self.generate_box_plot, width=button_width)
        btn_generate_box_plot.pack(padx=(0, 5), pady=(0, 5))
        btn_generate_box_plot.configure(style="Small.TButton")
        # --- Generate Priority-Category Bar Plot Button ---
        btn_priority_category_bar = ttk.Button(
            self.frame_top,
            text="Generate Priority-Category Bar Plot",
            command=self.generate_priority_category_bar_plot,
            width=button_width
        )
        btn_priority_category_bar.pack(padx=(0, 5), pady=(0, 5))
        btn_priority_category_bar.configure(style="Small.TButton")
        # --- Generate IssueType-Category Bar Plot Button ---
        btn_issuetype_category_bar = ttk.Button(
            self.frame_top,
            text="Generate IssueType-Category Bar Plot",
            command=self.generate_issuetype_category_bar_plot,
            width=button_width
        )
        btn_issuetype_category_bar.pack(padx=(0, 5), pady=(0, 5))
        btn_issuetype_category_bar.configure(style="Small.TButton")
        # --- Generate Word Count Button ---
        btn_wordcount = ttk.Button(self.frame_top, text="Generate Word Count Table", command=self.generate_wordcount_table, width=button_width)
        btn_wordcount.pack(padx=(0, 5), pady=(0, 5))
        btn_wordcount.configure(style="Small.TButton")
        # --- Info Label ---
        self.selected_info_label = tk.Label(self.frame_top, text="", bg=label_bg, font=label_font)
        self.selected_info_label.pack(fill='x', padx=10, pady=(0, 5))
        # --- Top N Variables ---
        defaultTopN = 10
        defaultTopNLabel = 'Top ' + str(defaultTopN)
        self.selected_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_top_n = tk.IntVar(value=defaultTopN)
        self.selected_wc_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_wc_top_n = tk.IntVar(value=defaultTopN)
        self.selected_bubble_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_bubble_top_n = tk.IntVar(value=defaultTopN)
        # --- Buttons and Menus ---
        self.setup_gui()

    def filter_defect_reports(self):
        from filter_defect_reports import filter_defect_reports_dialog
        filter_defect_reports_dialog(self)

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

    def generate_box_plot(self):
        file_path = filedialog.askopenfilename(
            title="Select Categorized Defect Report for Box Plot",
            filetypes=[("Excel files", "*categorized.xlsx;*categorized.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "No categorized defect report selected for box plot.")
            return
        import pandas as pd
        try:
            excel_data = pd.read_excel(file_path, sheet_name=None)
            priorities = set()
            has_priority = False
            for df in excel_data.values():
                df.columns = [col.lower() for col in df.columns]
                if "priority" in df.columns:
                    has_priority = True
                    priorities.update(str(p) for p in df["priority"].dropna().unique())
            priorities = sorted(priorities)
        except Exception:
            priorities = []
            has_priority = False
        selected_priorities = []
        if has_priority and priorities:
            # Multi-select dialog using checkboxes
            def show_priority_selector(priorities):
                dialog = tk.Toplevel(self)
                dialog.title("Select Priorities")
                dialog.grab_set()
                dialog.geometry("300x400")
                label = tk.Label(dialog, text="Select which priority to show:", font=label_font)
                label.pack(pady=10)
                var_dict = {}
                for p in priorities:
                    var = tk.BooleanVar(value=False)
                    cb = tk.Checkbutton(dialog, text=p, variable=var, font=small_button_font)
                    cb.pack(anchor='w', padx=20)
                    var_dict[p] = var
                def on_ok():
                    selected = [p for p, v in var_dict.items() if v.get()]
                    if not selected:
                        messagebox.showwarning("No Priority Selected", "Please select at least one priority.", parent=dialog)
                        return  # Do not close dialog
                    dialog.selected = selected # Always a list
                    dialog.destroy()
                def on_close():
                    dialog.selected = None
                    dialog.destroy()
                dialog.protocol("WM_DELETE_WINDOW", on_close)
                btn_all = tk.Button(dialog, text="Select All", command=lambda: [v.set(True) for v in var_dict.values()])
                btn_all.pack(pady=(10,0))
                btn_none = tk.Button(dialog, text="Deselect All", command=lambda: [v.set(False) for v in var_dict.values()])
                btn_none.pack()
                btn_ok = tk.Button(dialog, text="OK", command=on_ok)
                btn_ok.pack(pady=20)
                dialog.selected = [] # Always a list
                self.wait_window(dialog)
                return dialog.selected
            selected_priorities = show_priority_selector(priorities)
            # If user cancels dialog or closes window, do nothing
            if selected_priorities is None or not selected_priorities:
                return
            # If all priorities are selected, treat as 'All' (do not add priorities to title or filename)
            if isinstance(selected_priorities, list) and set(selected_priorities) == set(priorities):
                selected_priorities = "All"
        else:
            # No priority column found, just generate box plot for all data
            selected_priorities = None
        try:
            from plots import generate_category_box_plot
            generate_category_box_plot(file_path, selected_priorities)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate box plot:\n{str(e)}")

    def generate_priority_category_bar_plot(self):
        file_path = filedialog.askopenfilename(
            title="Select Categorized Defect Report for Bar Plot",
            filetypes=[("Excel files", "*categorized.xlsx;*categorized.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "No categorized defect report selected for bar plot.")
            return
        try:
            from plots import generate_priority_category_bar_plot
            generate_priority_category_bar_plot(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate bar plot:\n{str(e)}")

    def generate_issuetype_category_bar_plot(self):
        file_path = filedialog.askopenfilename(
            title="Select Categorized Defect Report for Bar Plot",
            filetypes=[("Excel files", "*categorized.xlsx;*categorized.xls")]
        )
        if not file_path:
            messagebox.showinfo("No Selection", "No categorized defect report selected for bar plot.")
            return
        try:
            from plots import generate_issue_type_category_bar_plot
            generate_issue_type_category_bar_plot(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate bar plot:\n{str(e)}")

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
            # Add 'All' option
            option_names = ['All Categories'] + sheet_names if len(sheet_names) > 1 else sheet_names
            # Remove previous OptionMenu if it exists
            if self.sheet_menu:
                self.sheet_menu.destroy()
                self.sheet_menu = None
            self.selected_sheet.set(option_names[0])
            self.sheet_menu = None
            if len(option_names) == 1:
                self.selected_info_label.config(
                    text=f"File: {os.path.basename(self.selected_excel_file)}"
                )
            else:
                self.sheet_menu = tk.OptionMenu(self.frame_top, self.selected_sheet, *option_names)
                self.sheet_menu.config(font=small_button_font, bg=btn_bg, highlightthickness=1, relief="groove")
                self.sheet_menu["menu"].config(font=small_button_font, bg=btn_bg)
                self.sheet_menu.pack(side='right', anchor='ne', padx=10, pady=(0, 10))
                self.selected_info_label.config(
                    text=f"File: {os.path.basename(self.selected_excel_file)}"
                )
                def update_label(*args):
                    selected = self.selected_sheet.get()
                    self.selected_info_label.config(
                        text=f"File: {os.path.basename(self.selected_excel_file)}"
                    )
                    try:
                        import pandas as pd
                        if selected == 'All Categories':
                            # Read all sheets and concatenate
                            all_dfs = []
                            excel_data = pd.read_excel(self.selected_excel_file, sheet_name=None)
                            for sheet, df in excel_data.items():
                                df.columns = [col.lower() for col in df.columns]
                                if "word" in df.columns and "count" in df.columns:
                                    all_dfs.append(df[[col for col in df.columns if col in ["word", "count"]]])
                            if all_dfs:
                                # Sort each df by descending count before concat
                                all_dfs = [df.sort_values('count', ascending=False) for df in all_dfs]
                                self.df = pd.concat(all_dfs, ignore_index=True)
                                self.df = self.df.sort_values('count', ascending=False).reset_index(drop=True)
                                self._missing_col_error_shown_per_sheet['All Categories'] = False
                            else:
                                if not self._missing_col_error_shown_per_sheet.get('All Categories', False):
                                    messagebox.showerror("Missing Columns", "None of the sheets contain both 'Word' and 'Count' columns.")
                                    self._missing_col_error_shown_per_sheet['All Categories'] = True
                                self.selected_info_label.config(text="")
                                self.df = None
                        else:
                            self.df = pd.read_excel(self.selected_excel_file, sheet_name=selected)
                            sheet = selected
                            self.df.columns = [col.lower() for col in self.df.columns]
                            if "word" not in self.df.columns or "count" not in self.df.columns:
                                if not self._missing_col_error_shown_per_sheet.get(sheet, False):
                                    messagebox.showerror("Missing Columns", "The selected file does not contain both 'Word' and 'Count' columns.")
                                    self._missing_col_error_shown_per_sheet[sheet] = True
                                self.selected_info_label.config(text="")
                                self.df = None
                            else:
                                self._missing_col_error_shown_per_sheet[sheet] = False
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to read sheet:\n{str(e)}")
                        self.df = None
                self.selected_sheet.trace_add('write', update_label)
            # Create dataframe for the first sheet by default
            import pandas as pd
            selected = self.selected_sheet.get()
            if selected == 'All Categories':
                all_dfs = []
                excel_data = pd.read_excel(self.selected_excel_file, sheet_name=None)
                for sheet, df in excel_data.items():
                    df.columns = [col.lower() for col in df.columns]
                    if "word" in df.columns and "count" in df.columns:
                        all_dfs.append(df[[col for col in df.columns if col in ["word", "count"]]])
                if all_dfs:
                    # Sort each df by descending count before concat
                    all_dfs = [df.sort_values('count', ascending=False) for df in all_dfs]
                    self.df = pd.concat(all_dfs, ignore_index=True)
                    self.df = self.df.sort_values('count', ascending=False).reset_index(drop=True)
                    self._missing_col_error_shown_per_sheet['All Categories'] = False
                else:
                    if not self._missing_col_error_shown_per_sheet.get('All Categories', False):
                        messagebox.showerror("Error", "None of the sheets contain both 'Word' and 'Count' columns.")
                        self._missing_col_error_shown_per_sheet['All Categories'] = True
                    self.selected_info_label.config(text="")
                    self.df = None
            else:
                self.df = pd.read_excel(self.selected_excel_file, sheet_name=selected)
                sheet = selected
                self.df.columns = [col.lower() for col in self.df.columns]
                if "word" not in self.df.columns or "count" not in self.df.columns:
                    if not self._missing_col_error_shown_per_sheet.get(sheet, False):
                        messagebox.showerror("Error", "There is no 'Word' or 'Count' column in the data")
                        self._missing_col_error_shown_per_sheet[sheet] = True
                    self.selected_info_label.config(text="")
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