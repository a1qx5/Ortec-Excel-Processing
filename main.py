from file_picker import pick_capacity_file, pick_actual_file, ask_output_file_name
from data_loader import load_capacity_file, load_actual_file
from data_processor import prepare_planned_df, prepare_actual_df, merge_and_process
from output_writer import save_final_excel


def main():
    # Pick files and output name
    capacity_file_path = pick_capacity_file()
    actual_file_path = pick_actual_file()
    output_file_name = ask_output_file_name()

    # Load and validate files
    capacity_sheets = load_capacity_file(capacity_file_path)
    actual_df = load_actual_file(actual_file_path)

    # Process data
    planned_df = prepare_planned_df(capacity_sheets)
    actual_long = prepare_actual_df(actual_df)
    merged_df = merge_and_process(planned_df, actual_long)

    # Save and style output
    save_final_excel(merged_df, output_file_name)

    print(f"Excel file '{output_file_name}' has been created successfully!")

if __name__ == "__main__":
    main()