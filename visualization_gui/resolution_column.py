import pandas as pd
from tkinter import messagebox
from openpyxl import load_workbook

def convert_to_date(value):
    try:
        return pd.to_datetime(value, errors="coerce", dayfirst=True)
    except:
        return value


def add_resolution_period_column_logic(file_path):
    # Load the Excel file
    original_wb = load_workbook(file_path)
    original_sheet_name = original_wb.sheetnames[0]

    # Read the sheet into a DataFrame
    df = pd.read_excel(file_path, engine="openpyxl", sheet_name=original_sheet_name)

    # Convert 'Created' and 'Resolved' columns to proper date formats
    if "Created" in df.columns:
        df["Created"] = df["Created"].apply(convert_to_date)
    if "Resolved" in df.columns:
        df["Resolved"] = df["Resolved"].apply(convert_to_date)

    # Add the 'Days spent to resolve' column
    if "Created" in df.columns and "Resolved" in df.columns:
        df["Days spent to resolve"] = (df["Resolved"] - df["Created"]).dt.days
    else:
        raise ValueError("Required columns 'Created' and 'Resolved' are missing.")

    # Save the updated DataFrame to a new Excel file
    output_excel = file_path.replace(".xlsx", "_with_resolution_period.xlsx")
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=original_sheet_name)

    # Save the workbook
    messagebox.showinfo("Success", f"Defect report with resolution period column saved to:\n{output_excel}")
