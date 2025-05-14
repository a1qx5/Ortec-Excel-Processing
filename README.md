# Automated Schedule Difference Reporter

## 📊 Overview

The Automated Schedule Difference Reporter is a Python utility designed to compare planned work capacity against actual hours logged. It processes two input Excel files—one detailing planned resource allocations and another containing actual time entries—and generates a consolidated Excel report showing the differences. This tool assists in tracking adherence to schedules and identifying variances in resource utilization.

The application uses a simple graphical user interface (GUI) for file selection and output naming.

## ✨ Features

* **GUI File Selection:** Provides dialogs for selecting input Excel files and specifying the output file name.
* **Data Aggregation:**
    * Processes planned capacity from multiple monthly sheets within a "Capacity Planning" Excel file.
    * Extracts and aggregates actual hours logged from an "Actual Hours" Excel file.
* **Monthly Comparison:** Aggregates both planned and actual data by employee, project, and month.
* **Difference Calculation:** Computes the variance between planned and actual hours for each entry.
* **Special Handling for Non-Working Time:** Identifies projects containing "vacation" or "leave" (case-insensitive) in their names and sets the calculated difference to `0` for these entries.
* **Formatted Excel Output:**
    * Generates a report in an `.xlsx` file named "Schedule Differences" by default (user can specify).
    * The output sheet includes **AutoFilter** dropdowns on column headers for easy sorting and filtering directly within Excel.
    * Highlights rows related to "vacation" or "leave" in orange for visual identification.
    * Sets a fixed width for the first six columns of the report for consistent readability.
* **Error Notifications:** Displays basic error messages via GUI dialogs for issues such as incorrect file format or inability to read files.

## ⚙️ Workflow

1.  **Run the Script:** Execute `main.py`.
2.  **Select Input Files:**
    * A dialog prompts for the "Capacity Planning" Excel file.
    * Another dialog prompts for the "Actual Hours" Excel file.
3.  **Specify Output File Name:** A dialog asks for the output report name (e.g., `Schedule_Differences.xlsx`).
4.  **Processing:** The script reads, cleans, transforms, and merges data from both input files.
5.  **Report Generation:** An Excel file is created with the comparison, showing planned hours, actual hours, their difference, and the applied formatting (highlighting, column widths, AutoFilter).
6.  **Confirmation:** A success message is displayed upon completion.

## 📁 Input File Requirements

The script requires two Excel (`.xlsx`) files:

### 1. Capacity Planning File

* **Content:** Should contain planned resource allocations, potentially across multiple sheets for different months.
* **Sheet Naming Convention:** Relevant sheets should contain `"Resource Requirements - "` in their names, followed by the month (e.g., `"Resource Requirements - Nov"`).
* **Expected Columns (within relevant sheets):**
    * `Resource`: Name of the employee/resource.
    * `Project`: Name of the project.
    * `Week X`: Columns for weekly planned days (e.g., `Week 44`, `Week 45`). Values are numeric (handles `2,5` or `2.5`).

### 2. Actual Hours File

* **Content:** Should contain actual hours logged by employees.
* **Sheet Name:** Data is read from a sheet named `"Hours"`.
* **Data Structure:** Data is expected to start from the third row (first two rows are skipped).
* **Expected Columns (column names are processed case-insensitively after stripping spaces):**
    * `employee`: Name of the employee.
    * `project`: Name of the project.
    * *Month Columns*: Columns named after months (e.g., `Oct`, `Nov`). Values are actual hours (numeric, handles `2,5` or `2.5`).

## 📋 Output File (`Schedule_Differences.xlsx`)

* **Sheet Name:** `"Schedule Differences"`
* **Key Columns:**
    * `Employee`
    * `Project`
    * `Month`
    * `Planned hours`
    * `Actual hours`
    * `Difference` (`Actual hours - Planned hours`; set to `0` for vacation/leave)
* **Excel Features:**
    * AutoFilter enabled on headers.
    * Relevant rows for "vacation" or "leave" projects are highlighted orange.
    * First six columns have a fixed width.

## 🚀 How to Run

There are two ways to use this application:

### 1. Using the Executable (Recommended for most users)

* Download the `ScheduleReporter.exe` (or similarly named `.exe`) file from the **[Releases Page](https://github.com/a1qx5/Ortec-Excel-Processing/releases)**.
* No installation is required. Double-click the `.exe` file.
* Follow the GUI prompts for file selection.
    * **Note:** The `.exe` bundles all necessary components.

### 2. Running from Python Source Code (For developers)

If you wish to run or modify the script from the source code:

#### Prerequisites (for running from source)

* Python 3.x (Developed and tested with Python 3.11.4)
* Required Python libraries: `pandas`, `openpyxl`.
    * `Tkinter` is also used (typically included with Python).

    Install libraries using pip:
    ```bash
    pip install pandas openpyxl
    ```

#### Running the Script (from source)

1.  Ensure Python 3 and the libraries are installed.
2.  Place all script files (`main.py`, `file_picker.py`, `data_loader.py`, `data_processor.py`, `output_writer.py`) in the same directory.
3.  Open a terminal or command prompt in that directory.
4.  Execute:
    ```bash
    python main.py
    ```
5.  Follow the GUI prompts.

## 🧩 Script Structure

* **`main.py`**: Orchestrates the overall workflow.
* **`file_picker.py`**: Manages GUI dialogs for file input/output.
* **`data_loader.py`**: Loads and validates data from the input Excel files.
* **`data_processor.py`**: Performs data cleaning, transformation, merging, and difference calculation.
* **`output_writer.py`**: Saves the final data to an Excel file and applies formatting (highlighting, column widths, AutoFilter).