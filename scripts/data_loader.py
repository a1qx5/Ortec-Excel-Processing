import pandas as pd
from tkinter import messagebox


def load_capacity_file(capacity_file_path):
    """Loads and checks the Capacity Planning Excel file."""
    try:
        test_capacity = pd.read_excel(capacity_file_path, sheet_name=None)

        # Find first correct sheet
        first_sheet_df = None
        for sheet_name, df in test_capacity.items():
            if "Resource Requirements - " in sheet_name:
                first_sheet_df = df
                break

        if first_sheet_df is None:
            messagebox.showerror("Error", "No valid sheet found in the Capacity Planning file!")
            exit()

        expected_columns = ['Resource', 'Project']
        if not all(col in first_sheet_df.columns for col in expected_columns):
            messagebox.showerror("Error", "The selected Capacity Planning file does not have the correct format!")
            exit()

        return test_capacity

    except Exception as e:
        messagebox.showerror("Error", f"Could not read Capacity Planning file: {e}")
        exit()

def load_actual_file(actual_file_path):
    """Loads and checks the Actual Hours Excel file."""
    try:
        actual_df = pd.read_excel(actual_file_path, sheet_name="Hours", skiprows=2)

        # Normalize columns to lowercase
        actual_df.columns = actual_df.columns.str.strip().str.lower()

        expected_columns = ['employee', 'project']
        if not all(col in actual_df.columns for col in expected_columns):
            messagebox.showerror("Error", "The selected Actual Hours file does not have the correct format!")
            exit()

        return actual_df

    except Exception as e:
        messagebox.showerror("Error", f"Could not read Actual Hours file: {e}")
        exit()