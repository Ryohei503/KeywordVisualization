import pandas as pd

def get_sheet_names(file_path):
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names

def get_save_path(base_path, sheet_names, sheet_name, suffix):
    # If only one sheet, do not include sheet name in file name
    if len(sheet_names) == 1:
        return f"{base_path}_{suffix}.png"
    else:
        return f"{base_path}_{sheet_name}_{suffix}.png"