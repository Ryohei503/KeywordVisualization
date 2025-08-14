import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from config import (
    primary_color, primary_dark, secondary_color,
    success_color, error_color, text_color, light_text, highlight_color,
    card_bg, border_color, header_font, button_font, 
    small_button_font, label_font, button_padding,
    frame_padding, section_padding, top_n_options
)
from get_utils import get_sheet_names, get_save_path


class VisualizationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Defect Report Analysis Tool")
        self.geometry("900x800")
        self.configure(bg=secondary_color)

        # Status bar
        self.status_bar = ttk.Label(
            self, 
            text="Ready", 
            relief="sunken", 
            anchor="w", 
            background=highlight_color,
            font=label_font,
            padding=(8, 4)
        )
        self.status_bar.pack(side="bottom", fill="x")
        
        # Configure ttk styles
        self.configure_styles()
        
        # --- Variables ---
        self.selected_excel_file = None
        self.selected_sheet = tk.StringVar(self)
        self.sheet_menu = None
        self.sheet_names_global = []
        
        # Create main container
        main_container = ttk.Frame(self, padding=frame_padding)
        main_container.pack(fill="both", expand=True)
        
        # Create left panel for analysis functions
        left_panel = ttk.LabelFrame(
            main_container, 
            text="Analysis Functions",
            padding=frame_padding,
            style="Card.TLabelframe"
        )
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10), ipadx=5, ipady=5)
        
        # Create right panel for visualization tools
        self.right_panel = ttk.LabelFrame(
            main_container, 
            text="Keyword Visualization",
            padding=frame_padding,
            style="Card.TLabelframe"
        )
        self.right_panel.pack(side="right", fill="both", expand=True, ipadx=5, ipady=5)
        
        # ===== LEFT PANEL CONTENT =====
        # Button groups with section headers
        sections = [
            ("Data Preparation", [
                ("Filter Defect Reports", self.filter_defect_reports),
                ("Add Resolution Period", self.add_resolution_period_column)
            ]),
            ("Categorization", [
                ("Build Categorization Model", self.build_model),
                ("Categorize Defect Reports", self.categorize_defects)
            ]),
            ("Categorized Data Plots", [
                ("Generate Category Pie Chart", self.generate_pie_chart),
                ("Generate Category Box Plot", self.generate_box_plot),
                ("Priority-Category Bar Plot", self.generate_priority_category_bar_plot),
                ("DefectType-Category Bar Plot", self.generate_defecttype_category_bar_plot)
            ]),
            ("Text Analysis", [
                ("Generate Word Count Table", self.generate_wordcount_table)
            ])
        ]
        
        for section_title, buttons in sections:
            # Section header
            header = ttk.Label(
                left_panel, 
                text=section_title, 
                style="SectionHeader.TLabel",
                padding=section_padding
            )
            header.pack(fill="x", pady=(10, 5))
            
            # Section buttons
            for btn_text, command in buttons:
                btn = ttk.Button(
                    left_panel, 
                    text=btn_text, 
                    command=command, 
                    style="Action.TButton",
                    width=25
                )
                btn.pack(fill="x", pady=3)

        # ===== RIGHT PANEL CONTENT =====
        # File selection section
        self.file_frame = ttk.LabelFrame(
            self.right_panel, 
            text="Word Count Table",
            style="Card.TLabelframe"
        )
        self.file_frame.pack(fill="x", pady=(0, 10))

        # File info
        self.selected_info_label = ttk.Label(
            self.file_frame, 
            text="No file selected", 
            font=label_font,
            background=highlight_color,
            padding=(8, 6),
            anchor="center"
        )
        self.selected_info_label.pack(fill="x", padx=5, pady=5)

        # File selection controls
        self.file_control_frame = ttk.Frame(self.file_frame)
        self.file_control_frame.pack(fill="x", pady=(0, 5))

        # Select file button
        self.btn_select_file = ttk.Button(
            self.file_control_frame, 
            text="Select Table", 
            command=self.select_excel_file, 
            style="Action.TButton"
        )
        self.btn_select_file.pack(side="left", padx=(0, 10))

        # Excel Sheet Data section (scrollable)
        self.sheet_data_frame = ttk.LabelFrame(
            self.file_frame,
            text="Excel Sheet Data (Top 100)",
            style="Card.TLabelframe"
        )
        self.sheet_data_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        # Create canvas and scrollbar for sheet data
        self.sheet_data_canvas = tk.Canvas(self.sheet_data_frame, bg="white", highlightthickness=0)
        self.sheet_data_scrollbar = ttk.Scrollbar(self.sheet_data_frame, orient="vertical", command=self.sheet_data_canvas.yview)
        self.sheet_data_canvas.configure(yscrollcommand=self.sheet_data_scrollbar.set)
        self.sheet_data_canvas.pack(side="left", fill="both", expand=True)
        self.sheet_data_scrollbar.pack(side="right", fill="y")

        # Frame inside canvas for table content
        self.sheet_data_inner = ttk.Frame(self.sheet_data_canvas)
        self.sheet_data_canvas.create_window((0, 0), window=self.sheet_data_inner, anchor="nw")

        # Bind resizing and scrolling
        self.sheet_data_inner.bind("<Configure>", lambda e: self.sheet_data_canvas.configure(scrollregion=self.sheet_data_canvas.bbox("all")))
        self.sheet_data_canvas.bind_all("<MouseWheel>", lambda e: self.sheet_data_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        vis_frame = ttk.LabelFrame(
        self.right_panel, 
        text="Visualization Options",
        style="Card.TLabelframe"
        )
        vis_frame.pack(fill="both", expand=True, pady=(10, 0))
    
        # Top N Variables
        defaultTopN = 10
        defaultTopNLabel = 'Top ' + str(defaultTopN)
        self.selected_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_top_n = tk.IntVar(value=defaultTopN)
        self.selected_wc_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_wc_top_n = tk.IntVar(value=defaultTopN)
        self.selected_bubble_top_n_label = tk.StringVar(value=defaultTopNLabel)
        self.selected_bubble_top_n = tk.IntVar(value=defaultTopN)

        # Visualization controls
        controls = [
            ("Graph", self.selected_top_n_label, self.on_generate_graph),
            ("Word Cloud", self.selected_wc_top_n_label, self.on_generate_wordcloud),
            ("Bubble Chart", self.selected_bubble_top_n_label, self.on_generate_bubble_chart)
        ]

        for i, (vis_type, var, command) in enumerate(controls):
            frame = ttk.Frame(vis_frame)
            frame.pack(fill="x", pady=5)

            # Visualization button
            btn = ttk.Button(
                frame, 
                text=f"Generate {vis_type}", 
                command=command, 
                style="Action.TButton",
                width=20
            )
            btn.pack(side="left", padx=(0, 10))

            # Top N label
            ttk.Label(frame, text="Show:", padding=(10, 0)).pack(side="left")
            
            # Top N menu
            menu_frame = ttk.Frame(frame)
            menu_frame.pack(side="left", padx=(0, 5))
            
            menu = tk.OptionMenu(
                menu_frame,
                var,
                *[label for label, val in top_n_options],
                command=lambda label, vis=vis_type: self.set_top_n(
                    dict(top_n_options)[label], 
                    label, 
                    vis.lower().replace(" ", "_")
                )
            )
            menu.config(
                font=small_button_font, 
                bg="white", 
                highlightthickness=1, 
                relief="groove", 
                width=10
            )
            menu["menu"].config(font=small_button_font, bg="white")
            menu.pack()
    
    def configure_styles(self):
        style = ttk.Style()
        
        # Configure main styles
        style.configure(".", background=secondary_color, foreground=text_color)
        style.configure("TFrame", background=secondary_color)
        style.configure("TLabel", background=secondary_color, foreground=text_color)
        style.configure("TLabelframe", background=secondary_color, relief="flat")
        style.configure("TLabelframe.Label", background=secondary_color, 
                        foreground=text_color, font=header_font)
        
        # Card style for frames
        style.configure(
        "Card.TLabelframe", 
        background=card_bg, 
        borderwidth=1,
        relief="solid",
        bordercolor=border_color,
        padding=frame_padding
    )
        
        # Section headers
        style.configure(
            "SectionHeader.TLabel",
            font=header_font,
            background=highlight_color,
            foreground=text_color,
            padding=section_padding,
            relief="flat",
            anchor="center"
        )
        
        # Action buttons (update to use border_color)
        style.configure(
        "Action.TButton",
        font=button_font,
        padding=button_padding,
        background="white",
        foreground=text_color,
        borderwidth=1,
        relief="solid",
        bordercolor=border_color
        )

        style.map(
            "Action.TButton",
            background=[("active", highlight_color)],
            relief=[("active", "sunken")]
        )
        
        # Primary buttons (right panel)
        style.configure(
            "Primary.TButton",
            font=button_font,
            padding=button_padding,
            background=primary_color,
            foreground=light_text,
            borderwidth=0,
            relief="flat",
            focuscolor=primary_dark
        )
        style.map(
            "Primary.TButton",
            background=[("active", primary_dark)],
            relief=[("active", "sunken")]
        )
        
        # Status colors
        style.configure("Success.TLabel", foreground=success_color)
        style.configure("Error.TLabel", foreground=error_color)

    # Update set_top_n methods to handle different visualization types
    def set_top_n(self, val, label, vis_type):
        if vis_type == "graph":
            self.selected_top_n.set(val)
            self.selected_top_n_label.set(label)
        elif vis_type == "word_cloud":
            self.selected_wc_top_n.set(val)
            self.selected_wc_top_n_label.set(label)
        elif vis_type == "bubble_chart":
            self.selected_bubble_top_n.set(val)
            self.selected_bubble_top_n_label.set(label)

    def filter_defect_reports(self):
        self.status_bar.config(text="Selecting defect report to filter...")
        file_path = filedialog.askopenfilename(
            title="Select a Defect Report to Filter", 
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if not file_path:
            self.status_bar.config(text="Operation canceled")
            return

        self.status_bar.config(text="Filtering a defect report...")
        from filter_defect_reports import filter_defect_reports_dialog
        success = filter_defect_reports_dialog(self, file_path)
        if success:
            self.status_bar.config(text="Defect report filtered successfully")
        else:
            self.status_bar.config(text="Failed to filter defect report")

    def add_resolution_period_column(self):
        self.status_bar.config(text="Selecting defect report to add resolution period...")
        file_path = filedialog.askopenfilename(
            title="Select a Defect Report", 
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if not file_path:
            self.status_bar.config(text="Operation canceled")
            return
            
        try:
            self.status_bar.config(text="Adding resolution period column...")
            from resolution_column import add_resolution_period_column_logic
            add_resolution_period_column_logic(file_path)
            self.status_bar.config(text="Resolution period added successfully")
        except Exception as e:
            self.status_bar.config(text="Failed to add resolution period")
            messagebox.showerror("Error", f"Failed to add resolution period column:\n{str(e)}")
    
    def build_model(self):
        self.status_bar.config(text="Selecting training data...")
        file_path = filedialog.askopenfilename(
            title="Select Train Data",
            filetypes=[("Excel files", "*.xlsx;*.xls")],
            parent=self
        )
        if not file_path:
            self.status_bar.config(text="Model training canceled")
            return

        # Progress dialog
        progress_win = tk.Toplevel(self)
        progress_win.title("Building Model")
        progress_win.geometry("300x150")
        progress_win.transient(self)
        progress_win.grab_set()
        
        # Add flag to track if training should be canceled
        self.cancel_training = False
        
        tk.Label(progress_win, text="Training model, please wait...", pady=10).pack()
        pb = ttk.Progressbar(progress_win, mode='indeterminate')
        pb.pack(fill='x', padx=20, pady=10)
        pb.start(10)
        
        # Add cancel button
        def cancel_training():
            self.cancel_training = True
            progress_win.destroy()
            self.status_bar.config(text="Model training canceled")
            messagebox.showinfo("Canceled", "Model training was canceled.", parent=self)
        
        cancel_btn = ttk.Button(
            progress_win, 
            text="Cancel", 
            command=cancel_training,
            style="Small.TButton"
        )
        cancel_btn.pack(pady=5)

        def worker():
            try:
                self.status_bar.config(text="Training model...")
                from build_model import train_model
                
                # Pass the cancel flag to the training function
                def should_cancel():
                    return self.cancel_training
                    
                success, msg = train_model(file_path, should_cancel)
                
                def on_done():
                    if progress_win.winfo_exists():  # Check if window still exists
                        pb.stop()
                        progress_win.destroy()
                    if success and not self.cancel_training:
                        self.status_bar.config(text="Model trained successfully")
                        messagebox.showinfo("Model Saved", msg, parent=self)
                    elif not self.cancel_training:
                        self.status_bar.config(text="Model training failed")
                        messagebox.showerror("Error", msg, parent=self)
                
                self.after(0, on_done)
            except Exception as e:
                def show_error():
                    if progress_win.winfo_exists():  # Check if window still exists
                        pb.stop()
                        progress_win.destroy()
                    if not self.cancel_training:
                        self.status_bar.config(text=f"Error: {str(e)}")
                        messagebox.showerror("Error", f"Model training failed:\n{str(e)}", parent=self)
                self.after(0, show_error)

        threading.Thread(target=worker, daemon=True).start()

    def categorize_defects(self):
        self.status_bar.config(text="Selecting defect report for categorization...")
        file_path = filedialog.askopenfilename(
            title="Select Defect Report to Categorize",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if not file_path:
            self.status_bar.config(text="Categorization canceled")
            return

        self.status_bar.config(text="Selecting model file...")
        model_path = filedialog.askopenfilename(
            title="Select Model File",
            filetypes=[("Pickle files", "*.pkl"), ("All files", "*.*")]
        )
        if not model_path:
            self.status_bar.config(text="Categorization canceled")
            messagebox.showinfo("No Selection", "No model selected for categorization.")
            return

        # Progress dialog starts after both files are selected
        progress_win = tk.Toplevel(self)
        progress_win.title("Categorizing Defect Report")
        progress_win.geometry("300x150")
        progress_win.transient(self)
        progress_win.grab_set()
        
        # Add flag to track if categorization should be canceled
        self.cancel_categorization = False
        
        tk.Label(progress_win, text="Categorizing defect report, please wait...", pady=10).pack()
        pb = ttk.Progressbar(progress_win, mode='indeterminate')
        pb.pack(fill='x', padx=20, pady=10)
        pb.start(10)

        # Add cancel button
        def cancel_categorization():
            self.cancel_categorization = True
            progress_win.destroy()
            self.status_bar.config(text="Categorization canceled")
            messagebox.showinfo("Canceled", "Categorization was canceled.", parent=self)
        
        cancel_btn = ttk.Button(
            progress_win, 
            text="Cancel", 
            command=cancel_categorization,
            style="Small.TButton"
        )
        cancel_btn.pack(pady=5)

        def worker():
            try:
                self.status_bar.config(text="Categorizing defect reports...")
                from summary_classifier import categorize_summaries
                
                # Pass the cancel flag to the categorization function
                def should_cancel():
                    return self.cancel_categorization
                    
                success = categorize_summaries(file_path, model_path, should_cancel)
                
                def on_done():
                    if progress_win.winfo_exists():  # Check if window still exists
                        pb.stop()
                        progress_win.destroy()
                    if success and not self.cancel_categorization:
                        self.status_bar.config(text="Categorization completed successfully")
                    elif not self.cancel_categorization:
                        self.status_bar.config(text="Categorization failed")
                
                self.after(0, on_done)
            except Exception as e:
                def show_error():
                    if progress_win.winfo_exists():  # Check if window still exists
                        pb.stop()
                        progress_win.destroy()
                    if not self.cancel_categorization:
                        self.status_bar.config(text=f"Error: {str(e)}")
                        messagebox.showerror("Error", f"Categorization failed:\n{str(e)}", parent=self)
                self.after(0, show_error)

        threading.Thread(target=worker, daemon=True).start()

    def generate_pie_chart(self):
        self.status_bar.config(text="Selecting categorized defect report...")
        file_path = filedialog.askopenfilename(
            title="Select Categorized Defect Report for Pie Chart",
            filetypes=[("Excel files", "*categorized.xlsx;*categorized.xls")]
        )
        if not file_path:
            self.status_bar.config(text="Pie chart generation canceled")
            return
        try:
            self.status_bar.config(text="Generating pie chart...")
            from plots import generate_category_pie_chart
            generate_category_pie_chart(file_path)
            self.status_bar.config(text="Pie chart generated successfully")
        except Exception as e:
            self.status_bar.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to generate pie chart:\n{str(e)}")

    def generate_box_plot(self):
        self.status_bar.config(text="Selecting categorized defect report...")
        file_path = filedialog.askopenfilename(
            title="Select Categorized Defect Report for Box Plot",
            filetypes=[("Excel files", "*categorized.xlsx;*categorized.xls")]
        )
        if not file_path:
            self.status_bar.config(text="Box plot generation canceled")
            return
            
        self.status_bar.config(text="Checking priorities...")
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
                self.status_bar.config(text="Box plot generation canceled")
                return
            # If all priorities are selected, treat as 'All' (do not add priorities to title or filename)
            if isinstance(selected_priorities, list) and set(selected_priorities) == set(priorities):
                selected_priorities = "All"
        else:
            # No priority column found, just generate box plot for all data
            selected_priorities = None
        
        self.status_bar.config(text="Generating box plot...")
        from plots import generate_category_box_plot
        success = generate_category_box_plot(file_path, selected_priorities)
        if success:
            self.status_bar.config(text="Box plot generated successfully")
        else:
            self.status_bar.config(text="Failed to generate box plot")

    def generate_priority_category_bar_plot(self):
        self.status_bar.config(text="Selecting categorized defect report...")
        file_path = filedialog.askopenfilename(
            title="Select Categorized Defect Report for Bar Plot",
            filetypes=[("Excel files", "*categorized.xlsx;*categorized.xls")]
        )
        if not file_path:
            self.status_bar.config(text="Bar plot generation canceled")
            return
    
        self.status_bar.config(text="Generating priority-category bar plot...")
        from plots import generate_priority_category_bar_plot
        success = generate_priority_category_bar_plot(file_path)
        if success:
            self.status_bar.config(text="Bar plot generated successfully")
        else:
            self.status_bar.config(text="Failed to generate bar plot")

    def generate_defecttype_category_bar_plot(self):
        self.status_bar.config(text="Selecting categorized defect report...")
        file_path = filedialog.askopenfilename(
            title="Select Categorized Defect Report for Bar Plot",
            filetypes=[("Excel files", "*categorized.xlsx;*categorized.xls")]
        )
        if not file_path:
            self.status_bar.config(text="Bar plot generation canceled")
            return
     
        self.status_bar.config(text="Generating defect type-category bar plot...")
        from plots import generate_defect_type_category_bar_plot
        success = generate_defect_type_category_bar_plot(file_path)
        if success:
            self.status_bar.config(text="Bar plot generated successfully")
        else:
            self.status_bar.config(text="Failed to generate bar plot")

    def generate_wordcount_table(self):
        self.status_bar.config(text="Selecting defect report for word count...")
        file_path = filedialog.askopenfilename(
            title="Select Defect Report for Word Count",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if not file_path:
            self.status_bar.config(text="Word count generation canceled")
            return
        self.status_bar.config(text="Filtering a defect report...")
        from word_count_util import word_count
        success = word_count(file_path)
        if success:
            self.status_bar.config(text="Word count table generated successfully")
        else:
            self.status_bar.config(text="Failed to generate a word count table")

    def select_excel_file(self):
        self.status_bar.config(text="Selecting word count table...")
        file_path = filedialog.askopenfilename(
            title="Select a Word Count Table",
            filetypes=[("Excel files", "*wordcount*.xlsx;*wordcount*.xls")]
        )
        if not file_path:
            self.status_bar.config(text="No file selected")
            return
            
        self.selected_excel_file = file_path
        try:
            self.status_bar.config(text="Loading sheet names...")
            sheet_names = get_sheet_names(self.selected_excel_file)
            self.sheet_names_global = sheet_names
            self._missing_col_error_shown_per_sheet = {}  # Track error per sheet

            import pandas as pd
            # Check first sheet for required columns before proceeding
            first_sheet = sheet_names[0] if sheet_names else None
            if first_sheet:
                df_first = pd.read_excel(self.selected_excel_file, sheet_name=first_sheet)
                df_first.columns = [col.lower() for col in df_first.columns]
                if "word" not in df_first.columns or "count" not in df_first.columns:
                    self.status_bar.config(text="File not loaded")
                    messagebox.showerror("Error", "The first sheet does not contain both 'Word' and 'Count' columns.")
                    self.df = None
                    return

            # Clear any existing sheet selection widgets first
            for widget in self.file_control_frame.winfo_children():
                if widget != self.btn_select_file:  # Keep the select button
                    widget.destroy()

            # Only show sheet selection if there are multiple sheets
            if len(sheet_names) > 1:
                # Add 'All' option
                option_names = ['All Categories'] + sheet_names

                # Remove previous OptionMenu if it exists
                if self.sheet_menu:
                    self.sheet_menu.destroy()
                    self.sheet_menu = None

                self.selected_sheet.set(option_names[0])

                # Add sheet selection controls
                ttk.Label(self.file_control_frame, text="Select sheet:").pack(side="left", padx=(10, 5))

                self.sheet_menu = tk.OptionMenu(
                    self.file_control_frame,
                    self.selected_sheet,
                    *option_names
                )
                self.sheet_menu.config(font=small_button_font, bg='white', 
                                    highlightthickness=1, relief="groove", width=15)
                self.sheet_menu["menu"].config(font=small_button_font, bg='white')
                self.sheet_menu.pack(side="left")

                # Set up trace for sheet selection
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
                                    all_dfs.append(df[["word", "count"]])
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
                                self.df = None
                        else:
                            self.df = pd.read_excel(self.selected_excel_file, sheet_name=selected)
                            sheet = selected
                            self.df.columns = [col.lower() for col in self.df.columns]
                            if "word" not in self.df.columns or "count" not in self.df.columns:
                                if not self._missing_col_error_shown_per_sheet.get(sheet, False):
                                    messagebox.showerror("Missing Columns", "The selected sheet does not contain both 'Word' and 'Count' columns.")
                                    self._missing_col_error_shown_per_sheet[sheet] = True
                                self.df = None
                            else:
                                self._missing_col_error_shown_per_sheet[sheet] = False
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to read sheet:\n{str(e)}")
                        self.df = None

                self.selected_sheet.trace_add('write', update_label)

                # Update sheet data display when sheet changes
                def update_sheet_data_and_display(*args):
                    update_label()
                    self.update_sheet_data_display()
                self.selected_sheet.trace_add('write', update_sheet_data_and_display)

            # Load data (for single sheet or default selection)
            if len(sheet_names) > 1:
                selected = self.selected_sheet.get()
                if selected == 'All Categories':
                    all_dfs = []
                    excel_data = pd.read_excel(self.selected_excel_file, sheet_name=None)
                    for sheet, df in excel_data.items():
                        df.columns = [col.lower() for col in df.columns]
                        if "word" in df.columns and "count" in df.columns:
                            all_dfs.append(df[["word", "count"]])
                    if all_dfs:
                        all_dfs = [df.sort_values('count', ascending=False) for df in all_dfs]
                        self.df = pd.concat(all_dfs, ignore_index=True)
                        self.df = self.df.sort_values('count', ascending=False).reset_index(drop=True)
                    else:
                        self.df = None
                else:
                    self.df = pd.read_excel(self.selected_excel_file, sheet_name=selected)
                    self.df.columns = [col.lower() for col in self.df.columns]
            else:
                # Single sheet - load it directly
                self.df = pd.read_excel(self.selected_excel_file)
                self.df.columns = [col.lower() for col in self.df.columns]

            # Validate columns
            if self.df is not None and ("word" not in self.df.columns or "count" not in self.df.columns):
                messagebox.showerror("Error", "The selected sheet does not contain both 'Word' and 'Count' columns.")
                self.df = None

            self.selected_info_label.config(
                text=f"File: {os.path.basename(file_path)}"
            )
            self.status_bar.config(text=f"Loaded: {os.path.basename(file_path)}")
            self.update_sheet_data_display()

        except Exception as e:
            self.status_bar.config(text="Failed to load file")
            messagebox.showerror("Error", f"Failed to read sheets:\n{str(e)}")

    def update_sheet_data_display(self):
        # Clear previous content
        for widget in self.sheet_data_inner.winfo_children():
            widget.destroy()
        
        # Display DataFrame if available
        import pandas as pd
        if hasattr(self, 'df') and isinstance(self.df, pd.DataFrame) and self.df is not None:
            df = self.df.copy()
            # Limit rows/cols for display if large
            max_rows = 100
            max_cols = 15
            display_df = df.head(max_rows)
            columns = list(display_df.columns)[:max_cols]
            
            # Header
            for j, col in enumerate(columns):
                lbl = ttk.Label(
                    self.sheet_data_inner, 
                    text=str(col), 
                    font=label_font, 
                    background=highlight_color, 
                    borderwidth=1, 
                    relief="solid", 
                    padding=2
                )
                lbl.grid(row=0, column=j, sticky="nsew")
            
            # Data rows
            for i, row in enumerate(display_df.itertuples(index=False), start=1):
                for j, val in enumerate(row[:max_cols]):
                    lbl = ttk.Label(
                        self.sheet_data_inner, 
                        text=str(val), 
                        font=small_button_font, 
                        background="white", 
                        borderwidth=1, 
                        relief="solid", 
                        padding=2
                    )
                    lbl.grid(row=i, column=j, sticky="nsew")
            
            # Make columns expand
            for j in range(len(columns)):
                self.sheet_data_inner.grid_columnconfigure(j, weight=1)
        else:
            lbl = ttk.Label(
                self.sheet_data_inner, 
                text="No data to display.", 
                font=label_font, 
                background="white"
            )
            lbl.grid(row=0, column=0, sticky="nsew")

    def on_generate_graph(self):
        if not self.selected_excel_file or self.df is None:
            self.status_bar.config(text="No word count table selected")
            messagebox.showwarning("No File Selected", "Please select a word count table first.")
            return
        try:
            self.status_bar.config(text="Loading fonts...")
            from plots import load_japanese_font, get_colordict, generate_graph
            font_path = load_japanese_font()
            if not font_path:
                self.status_bar.config(text="Font loading failed")
                messagebox.showerror("Error", "Failed to read fonts")
                return
                
            n = self.selected_top_n.get()
            actual_n = min(n, len(self.df))
            if len(self.df) < n:
                messagebox.showwarning(
                    "Not Enough Words",
                    f"The selected word count table does not have enough words. Showing {actual_n} words instead of {n}."
                )
                
            self.status_bar.config(text="Generating graph...")
            df_top = self.df.iloc[:actual_n]
            chunk_size = 20
            max_count = df_top['count'].max()
            color_dict = get_colordict('viridis', max_count, 1)
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            top_n_label = f"Top{actual_n}" if actual_n != n else self.selected_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{top_n_label}_graph")
            generate_graph(df_top, chunk_size, color_dict, f"{top_n_label}", sheet_name, output_path, self.sheet_names_global)
            self.status_bar.config(text="Graph generated successfully")
            messagebox.showinfo("Completed", f"Saved a graph:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            self.status_bar.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"There was an error:\n{str(e)}")

    def on_generate_wordcloud(self):
        if not self.selected_excel_file or self.df is None:
            self.status_bar.config(text="No word count table selected")
            messagebox.showwarning("No File Selected", "Please select a word count table first.")
            return
        try:
            self.status_bar.config(text="Loading fonts...")
            from plots import load_japanese_font, generate_wordcloud
            font_path = load_japanese_font()
            if not font_path:
                self.status_bar.config(text="Font loading failed")
                messagebox.showerror("Error", "Failed to read fonts")
                return
                
            n = self.selected_wc_top_n.get()
            actual_n = min(n, len(self.df))
            if len(self.df) < n:
                messagebox.showwarning(
                    "Not Enough Words",
                    f"The selected word count table does not have enough words. Showing {actual_n} words instead of {n}."
                )
                
            self.status_bar.config(text="Generating word cloud...")
            wc_top_n_label = f"Top{actual_n}" if actual_n != n else self.selected_wc_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{wc_top_n_label}_wordcloud")
            generate_wordcloud(self.df.iloc[:actual_n], actual_n, font_path, output_path, sheet_name, wc_top_n_label, self.sheet_names_global)
            self.status_bar.config(text="Word cloud generated successfully")
            messagebox.showinfo("Completed", f"Saved a word cloud:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            self.status_bar.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"There was an error:\n{str(e)}")

    def on_generate_bubble_chart(self):
        if not self.selected_excel_file or self.df is None:
            self.status_bar.config(text="No word count table selected")
            messagebox.showwarning("No File Selected", "Please select a word count table first.")
            return
        try:
            self.status_bar.config(text="Loading fonts...")
            from plots import load_japanese_font, get_colordict, generate_bubble_chart
            font_path = load_japanese_font()
            if not font_path:
                self.status_bar.config(text="Font loading failed")
                messagebox.showerror("Error", "Failed to read fonts")
                return
                
            n = self.selected_bubble_top_n.get()
            actual_n = min(n, len(self.df))
            if len(self.df) < n:
                messagebox.showwarning(
                    "Not Enough Words",
                    f"The selected word count table does not have enough words. Showing {actual_n} words instead of {n}."
                )
                
            self.status_bar.config(text="Generating bubble chart...")
            bubble_top_n_label = f"Top{actual_n}" if actual_n != n else self.selected_bubble_top_n_label.get().replace(" ", "")
            base_path = os.path.splitext(self.selected_excel_file)[0]
            sheet_name = self.selected_sheet.get().replace(" ", "_")
            max_count = self.df['count'].max()
            color_dict = get_colordict('summer', max_count, min(self.df['count']))
            output_path = get_save_path(base_path, self.sheet_names_global, sheet_name, f"{bubble_top_n_label}_bubblechart")
            generate_bubble_chart(self.df.iloc[:actual_n], actual_n, color_dict, output_path, sheet_name, bubble_top_n_label,self.sheet_names_global)
            self.status_bar.config(text="Bubble chart generated successfully")
            messagebox.showinfo("Completed", f"Saved a bubble chart:\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            self.status_bar.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"There was an error:\n{str(e)}")
