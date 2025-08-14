import pandas as pd
from tkinter import filedialog, messagebox

def filter_defect_reports_dialog(parent):
    input_excel = filedialog.askopenfilename(
        title="Select a Defect Report to Filter",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    if not input_excel:
        messagebox.showinfo("No Selection", "No file selected.")
        return
    try:
        excel_data = pd.read_excel(input_excel, sheet_name=None)
        found_columns = set()
        for df in excel_data.values():
            df.columns = [col.lower() for col in df.columns]
            if "priority" in df.columns:
                found_columns.add("Priority")
            if "custom field (category)" in df.columns:
                found_columns.add("Custom field (Category)")
        if not found_columns:
            messagebox.showerror("No Filter Column", "There is no column to filter the defect reports (must have 'Priority', 'Custom field (Category)', 'Created', or 'Resolved').")
            return
        # Multi-column filter dialog for 'Priority', 'Custom field (Category)', 'created', and 'resolved'
        import tkinter as tk
        from datetime import datetime
        # Gather all unique values for each categorical filter column
        filter_columns = [col for col in ["Priority", "Custom field (Category)"] if col in found_columns]
        date_columns = [col for col in ["created", "resolved"] if any(col in df.columns for df in excel_data.values())]
        unique_values_dict = {}
        for col in filter_columns:
            col_key = col.lower()
            values = set()
            for df in excel_data.values():
                df.columns = [c.lower() for c in df.columns]
                if col_key in df.columns:
                    values.update(df[col_key].dropna().unique())
            unique_values_dict[col] = sorted(values)

        class MultiFilterDialog(tk.Toplevel):
            def __init__(self, parent, columns_values, date_columns):
                super().__init__(parent)
                self.title("Select Filter Values")
                self.selected = {}
                self.vars = {}
                self.columns_values = columns_values  # Save for use in on_ok
                tk.Label(self, text="Select values to filter the defect reports:").pack(padx=10, pady=5)
                # Categorical columns
                for col, options in columns_values.items():
                    frame = tk.LabelFrame(self, text=col)
                    frame.pack(padx=10, pady=5, fill='x')
                    self.vars[col] = []
                    for val in options:
                        var = tk.BooleanVar()
                        cb = tk.Checkbutton(frame, text=str(val), variable=var)
                        cb.pack(anchor='w')
                        self.vars[col].append((var, val))
                # Date columns
                self.date_entries = {}
                for col in date_columns:
                    frame = tk.LabelFrame(self, text=col.capitalize() + " (YYYY-MM-DD)")
                    frame.pack(padx=10, pady=5, fill='x')
                    tk.Label(frame, text="From:").pack(side='left', padx=2)
                    from_entry = tk.Entry(frame, width=12)
                    from_entry.pack(side='left', padx=2)
                    tk.Label(frame, text="To:").pack(side='left', padx=2)
                    to_entry = tk.Entry(frame, width=12)
                    to_entry.pack(side='left', padx=2)
                    self.date_entries[col] = (from_entry, to_entry)
                btn_frame = tk.Frame(self)
                btn_frame.pack(pady=5)
                tk.Button(btn_frame, text="OK", command=self.on_ok).pack(side='left', padx=5)
                tk.Button(btn_frame, text="Cancel", command=self.on_cancel).pack(side='left', padx=5)
                self.protocol("WM_DELETE_WINDOW", self.on_cancel)
                self.transient(parent)
                self.grab_set()
                self.wait_window(self)
            def on_ok(self):
                for col, varlist in self.vars.items():
                    selected = [val for var, val in varlist if var.get()]
                    self.selected[col] = selected
                # Date ranges
                self.selected['date_ranges'] = {}
                for col, (from_entry, to_entry) in self.date_entries.items():
                    from_val = from_entry.get().strip()
                    to_val = to_entry.get().strip()
                    self.selected['date_ranges'][col] = (from_val, to_val)
                # Require at least one value selected for at least one column or a date range
                has_cat = any(self.selected.get(col) for col in self.columns_values)
                has_date = any(self.selected['date_ranges'][col][0] or self.selected['date_ranges'][col][1] for col in self.date_entries)
                if not (has_cat or has_date):
                    messagebox.showwarning("Selection Required", "You must select at least one value or enter a date range to filter.", parent=self)
                    return
                self.destroy()
            def on_cancel(self):
                self.selected = None
                self.destroy()

        dialog = MultiFilterDialog(parent, unique_values_dict, date_columns)
        selected_values_dict = dialog.selected
        if not selected_values_dict:
            return

        # Filter each sheet and collect results
        filtered_dfs = {}
        for sheet_name, df in excel_data.items():
            df.columns = [col.lower() for col in df.columns]
            mask = pd.Series([True] * len(df))
            # Categorical filters
            for col in filter_columns:
                col_key = col.lower()
                selected_vals = selected_values_dict.get(col, [])
                if selected_vals and col_key in df.columns:
                    mask &= df[col_key].isin(selected_vals)
            # Date filters
            date_ranges = selected_values_dict.get('date_ranges', {})
            for col in date_columns:
                col_key = col.lower()
                if col_key in df.columns:
                    from_str, to_str = date_ranges.get(col, ('', ''))
                    if from_str:
                        try:
                            from_date = pd.to_datetime(from_str, errors='coerce', dayfirst=True)
                            mask &= pd.to_datetime(df[col_key], errors='coerce', dayfirst=True) >= from_date
                        except Exception:
                            pass
                    if to_str:
                        try:
                            to_date = pd.to_datetime(to_str, errors='coerce', dayfirst=True)
                            mask &= pd.to_datetime(df[col_key], errors='coerce', dayfirst=True) <= to_date
                        except Exception:
                            pass
            filtered = df[mask]
            if not filtered.empty:
                filtered_dfs[sheet_name] = filtered
        if not filtered_dfs:
            messagebox.showinfo("No Results", "No defect reports found with the selected filters.")
            return
        # Save filtered results to a new Excel file, ensuring date columns are datetime
        import os
        output_path = input_excel.replace(".xlsx", "_filtered.xlsx")
        date_cols = [col for col in ["created", "resolved"] if any(col in df.columns for df in filtered_dfs.values())]
        with pd.ExcelWriter(output_path, date_format="yyyy-mm-dd HH:MM:SS", datetime_format="yyyy-mm-dd HH:MM:SS") as writer:
            for sheet_name, df in filtered_dfs.items():
                # Convert date columns to datetime if present
                for dcol in date_cols:
                    if dcol in df.columns:
                        df.loc[:, dcol] = pd.to_datetime(df[dcol], errors='coerce', dayfirst=True)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        messagebox.showinfo("Filtered Results Saved", f"Filtered defect reports saved to:\n{output_path}")
        try:
            os.startfile(output_path)
        except Exception:
            pass
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read defect report:\n{str(e)}")
