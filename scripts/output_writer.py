import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from tkinter import messagebox, filedialog


def save_final_excel(merged_df):
    """Save the merged DataFrame to Excel."""
    output_file_path = filedialog.asksaveasfilename(title="Save Excel file as",
                                                    defaultextension=".xlsx",
                                                    filetypes=[("Excel files", "*.xlsx")],
                                                    initialfile="Schedule_Differences.xlsx"
                                                    )

    if not output_file_path:
        messagebox.showinfo("Cancelled", "Save cancelled by user.")
        exit()

    merged_df.to_excel(output_file_path, index=False, sheet_name="Schedule Differences")
    highlight_vacation_rows(output_file_path)

    messagebox.showinfo("Saved", f"Excel file successfully saved as: \n{output_file_path}")
def highlight_vacation_rows(file_path):
    """Highlight vacation/leave rows in orange (planned, actual, diff columns)."""
    wb = load_workbook(file_path)
    ws = wb.active

    # Define orange fill color
    orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

    # Get column positions
    columns = {cell.value: idx for idx, cell in enumerate(ws[1], 1)}  # header row is row 1

    # Loop over rows and highlight cells for vacation/leave projects
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        project_value = str(row[columns['project'] - 1].value).lower()
        if 'vacation' in project_value or 'leave' in project_value:
            for col in ['planned_hours', 'actual_hours', 'diff', 'project', 'month', 'employee']:
                col_idx = columns[col] - 1
                row[col_idx].fill = orange_fill

    wb.save(file_path)