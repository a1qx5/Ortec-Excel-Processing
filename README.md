# Automated Schedule Difference Reporter

## 📊 Overview

The Automated Schedule Difference Reporter is a Python-based utility designed to streamline the process of comparing planned work capacity against actual hours logged. It takes two Excel files as input—one detailing planned resource allocation and another containing actual time entries—processes the data, and generates a consolidated Excel report highlighting the differences. This tool is particularly useful for project managers and teams looking to track adherence to schedules and identify discrepancies in resource utilization.

The application features a simple graphical user interface (GUI) for easy file selection and output naming, making it accessible even for users with limited technical background.

## ✨ Features

* **GUI File Selection:** User-friendly dialogs for selecting input Excel files and specifying the output file name.
* **Automated Data Aggregation:** Processes planned capacity from multiple monthly sheets within the Capacity Planning Excel file.
* **Monthly Comparison:** Aggregates planned data by employee, project, and month.
* **Actual Hours Processing:** Extracts and aggregates actual hours logged by employee, project, and month.
* **Difference Calculation:** Computes the variance between planned and actual hours.
* **Special Handling for Non-Working Time:** Identifies projects containing "vacation" or "leave" and neutralizes the difference calculation for these entries.
* **Formatted Excel Output:**
    * Generates a clean, consolidated report in an `.xlsx` file.
    * Highlights rows related to "vacation" or "leave" in orange for easy visual identification.
    * Sets a fixed width for the first six columns of the report for consistent readability.
* **Error Notifications:** Provides basic error messages via GUI dialogs for issues like incorrect file format or inability to read files.

## ⚙️ Workflow

1.  **Run the Script:** Execute the `main.py` script.
2.  **Select Capacity Planning File:** A dialog will prompt you to select your "Capacity Planning" Excel file.
3.  **Select Actual Hours File:** Next, a dialog will prompt you to select your "Actual Hours" Excel file.
4.  **Specify Output File Name:** You'll be asked to enter a name for the output report (defaults to `Schedule_Differences.xlsx`).
5.  **Processing:** The script reads, cleans, transforms, and merges the data from both input files.
6.  **Report Generation:** An Excel file is generated containing the comparison, with planned hours, actual hours, and their difference, along with formatting for vacation/leave and column widths.
7.  **Confirmation:** A success message is displayed upon completion.

## 📁 Input File Requirements

The script expects two specific Excel (`.xlsx`) files as input:

### 1. Capacity Planning File

* **Content:** This file should contain planned resource allocations, typically with separate sheets for different months.
* **Sheet Naming Convention:** The script looks for sheets whose names contain `"Resource Requirements - "`, followed by the month name (e.g., `"Resource Requirements - Nov"`, `"Resource Requirements - Jan"`).
* **Expected Columns (within each relevant sheet):**
    * `Resource`: Name of the employee or resource.
    * `Project`: Name of the project.
    * `Week X`: Multiple columns representing weeks (e.g., `Week 44`, `Week 45`). Values under these columns should be planned days (numeric, decimals like `2,5` or `2.5` are handled).

### 2. Actual Hours File

* **Content:** This file should contain the actual hours logged by employees on various projects.
* **Sheet Name:** The script specifically reads data from a sheet named `"Hours"`.
* **Data Structure:** The script expects the actual data to start from the third row (it skips the first two rows).
* **Expected Columns (case-insensitive after stripping spaces):**
    * `employee`: Name of the employee.
    * `project`: Name of the project.
    * *Month Columns*: Columns named after months (e.g., `Oct`, `Nov`, `Dec`). Values under these columns should be actual hours logged (numeric, decimals like `2,5` or `2.5` are handled).

## 📋 Output File

* **File Name:** User-defined via a save dialog, defaulting to `Schedule_Differences.xlsx`.
* **Sheet Name:** The report is generated in a sheet named `"Schedule Differences"`.
* **Key Columns in the Output:**
    * `Employee`: Name of the employee.
    * `Project`: Name of the project.
    * `Month`: The month of the record.
    * `Planned hours`: Total planned hours for the employee on the project for that month.
    * `Actual hours`: Total actual hours logged by the employee on the project for that month.
    * `Difference`: Calculated as `Actual hours - Planned hours`. For "vacation" or "leave" projects, this is set to `0`.
* **Formatting:**
    * Rows where the 'Project' column contains "vacation" or "leave" (case-insensitive) will have the 'Planned hours', 'Actual hours', 'Difference', 'Project', 'Month', and 'Employee' cells highlighted in orange.
    * The first six columns of the report will have a fixed width (approximately 15 character units) for better readability.

## 🚀 How to Run

There are two ways to use this application:

### 1. Using the Executable (Recommended for most users)

* Download the `ScheduleReporter.exe` file from the **[Releases](https://github.com/a1qx5/Ortec-Excel-Processing/releases)**.
* No installation is required. Simply double-click the `.exe` file to run it.
* Follow the on-screen GUI prompts to select your input files and specify the output file name.
    * **Note:** The `.exe` includes all necessary components and does not require a separate Python installation.

### 2. Running from Python Source Code (For developers)

If you wish to run the script directly from the source code or modify it:

#### Prerequisites (for running from source)

* Python 3.x (tested with Python 3.x.x)
* The following Python libraries:
    * `pandas`
    * `openpyxl`
    * `Tkinter` (usually included with standard Python installations, but ensure your Python environment has it)

    You can install the necessary libraries using pip:
    ```bash
    pip install pandas openpyxl
    ```

#### Running the Script (from source)

1.  Ensure you have Python 3 and the required libraries installed (see prerequisites above).
2.  Download or clone all the script files (`main.py`, `file_picker.py`, `data_loader.py`, `data_processor.py`, `output_writer.py`) into the same directory.
3.  Open a terminal or command prompt.
4.  Navigate to the directory where you saved the files.
5.  Run the main script using the command:
    ```bash
    python main.py
    ```
6.  Follow the on-screen GUI prompts.