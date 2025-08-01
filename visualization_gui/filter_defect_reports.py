import pandas as pd
from tkinter import filedialog, messagebox
import tkinter.simpledialog as simpledialog

def filter_defect_reports_dialog(parent):
    file_path = filedialog.askopenfilename(
        title="Select a Defect Report to Filter",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    if not file_path:
        messagebox.showinfo("No Selection", "No file selected.")
        return
    try:
        excel_data = pd.read_excel(file_path, sheet_name=None)
        found_columns = set()
        for df in excel_data.values():
            df.columns = [col.lower() for col in df.columns]
            if "priority" in df.columns:
                found_columns.add("Priority")
            if "custom field (category)" in df.columns:
                found_columns.add("Custom field (Category)")
        if not found_columns:
            messagebox.showerror("No Filter Column", "There is no column to filter the defect reports (must have 'Priority' or 'Custom field (Category)').")
            return
        filter_options = list(found_columns)
        selected_filter = simpledialog.askstring(
            "Select Filter Column",
            f"Select how you want to filter the defect reports:\nOptions: {', '.join(filter_options)}\n(Type exactly as shown)",
            initialvalue=filter_options[0]
        )
        if selected_filter is None or selected_filter not in filter_options:
            return
        messagebox.showinfo("Selected Filter", f"You selected to filter by: {selected_filter}")
        # Ask user for filter value
        filter_value = simpledialog.askstring(
            "Filter Value",
            f"Enter the value to filter '{selected_filter}' by:",
            initialvalue=""
        )
        if filter_value is None or filter_value.strip() == "":
            return
        # Filter each sheet and collect results
        filtered_dfs = {}
        for sheet_name, df in excel_data.items():
            df.columns = [col.lower() for col in df.columns]
            col_key = selected_filter.lower()
            if col_key in df.columns:
                filtered = df[df[col_key] == filter_value]
                if not filtered.empty:
                    filtered_dfs[sheet_name] = filtered
        if not filtered_dfs:
            messagebox.showinfo("No Results", f"No defect reports found with {selected_filter} = '{filter_value}'.")
            return
        # Save filtered results to a new Excel file
        import os
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_filtered_{selected_filter}_{filter_value}{ext}"
        with pd.ExcelWriter(output_path) as writer:
            for sheet_name, df in filtered_dfs.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        messagebox.showinfo("Filtered Results Saved", f"Filtered defect reports saved to:\n{output_path}")
        try:
            os.startfile(output_path)
        except Exception:
            pass
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read defect report:\n{str(e)}")
