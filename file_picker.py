from tkinter import Tk, filedialog, simpledialog


def pick_capacity_file():
    """Open a file picker to select the Capacity Planning Excel file."""
    Tk().withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Capacity Planning Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    return file_path

def pick_actual_file():
    """Open a file picker to select the Actual Hours Excel file."""
    Tk().withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Actual Hours Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    return file_path

def ask_output_file_name():
    """Ask user for output file name and add .xlsx if missing."""
    Tk().withdraw()
    file_name = simpledialog.askstring("Save As", "Enter the output Excel file name:")
    if not file_name.endswith(".xlsx"):
        file_name += ".xlsx"
    return file_name