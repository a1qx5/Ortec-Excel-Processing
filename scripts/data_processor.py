import pandas as pd


def process_capacity_sheet(df, month_name):
    """Processes one capacity planning sheet."""
    df.columns = [str(col).strip() for col in df.columns]

    id_cols = [col for col in df.columns if col in ['Resource', 'Project']]
    week_cols = [col for col in df.columns if 'Week' in col]

    long_df = df.melt(id_vars=id_cols, value_vars=week_cols,
                      var_name='week', value_name='planned_days')

    long_df.dropna(subset=['planned_days'], inplace=True)

    long_df['planned_days'] = long_df['planned_days'].astype(str).str.replace(',', '.', regex=False)
    long_df['planned_days'] = pd.to_numeric(long_df['planned_days'], errors='coerce')
    long_df.dropna(subset=['planned_days'], inplace=True)

    long_df['month'] = month_name
    long_df['planned_hours'] = long_df['planned_days'] * 8

    result = long_df.groupby(['Resource', 'Project', 'month'], as_index=False)['planned_hours'].sum()
    result.rename(columns={'Resource': 'employee', 'Project': 'project'}, inplace=True)
    result['month'] = result['month'].astype(str)

    return result

def prepare_planned_df(capacity_sheets):
    """Processes all sheets into a full planned DataFrame."""
    all_months = []

    for sheet_name, df in capacity_sheets.items():
        if "Resource Requirements - " in sheet_name:
            month = sheet_name.split(" - ")[-1]
            processed = process_capacity_sheet(df, month)
            all_months.append(processed)

    planned_df = pd.concat(all_months, ignore_index=True)
    # Ensure month is always string type
    planned_df['month'] = planned_df['month'].astype(str)
    return planned_df

def prepare_actual_df(actual_df):
    """Processes the actual worked hours DataFrame."""
    actual_df.columns = actual_df.columns.map(str)

    month_columns = [col for col in actual_df.columns if col not in ['employee', 'project', 'total']]

    actual_df.columns = actual_df.columns.str.strip().str.lower()

    actual_long = actual_df.melt(id_vars=['employee', 'project'],
                                 value_vars=month_columns,
                                 var_name='month', value_name='actual_hours')

    actual_long['month'] = actual_long['month'].astype(str)
    actual_long['month'] = actual_long['month'].str.extract(r'([A-Za-z]+)')
    actual_long.dropna(subset=['month'], inplace=True)
    actual_long['month'] = actual_long['month'].str.capitalize()

    actual_long['actual_hours'] = actual_long['actual_hours'].astype(str).str.replace(',', '.', regex=False)
    actual_long['actual_hours'] = pd.to_numeric(actual_long['actual_hours'], errors='coerce')
    actual_long.dropna(subset=['actual_hours'], inplace=True)

    actual_long = actual_long.groupby(['employee', 'project', 'month'], as_index=False)['actual_hours'].sum()

    # Ensure month is always string type
    actual_long['month'] = actual_long['month'].astype(str)
    return actual_long

def merge_and_process(planned_df, actual_long):
    """Merges planned and actuals, calculates diffs, handles vacations."""

    # Ensure consistent string types for all merge columns right before merging
    merge_cols = ['employee', 'project', 'month']
    for col in merge_cols:
        if col in planned_df.columns:
            planned_df[col] = planned_df[col].astype(str)
        if col in actual_long.columns:
            actual_long[col] = actual_long[col].astype(str)

    # Perform the merge
    merged_df = pd.merge(planned_df, actual_long, on=merge_cols, how='outer')

    # Fill NaN values introduced by the outer merge in numeric columns
    merged_df['planned_hours'] = merged_df['planned_hours'].fillna(0)
    merged_df['actual_hours'] = merged_df['actual_hours'].fillna(0)

    # Calculate difference
    merged_df['Difference'] = merged_df['actual_hours'] - merged_df['planned_hours']

    # Handle Vacation/Leave rows - ensure 'project' column is string for .str accessor
    merged_df['project'] = merged_df['project'].astype(str) # Ensure project is string type after merge
    vacation_mask = merged_df['project'].str.lower().str.contains('vacation|leave', na=False)
    merged_df.loc[vacation_mask, 'Difference'] = 0 # Set Difference to 0 for vacation/leave

    # Final sorting
    merged_df.sort_values(by=['employee', 'month', 'project'], inplace=True)

    # Capitalize & remove _ from columns
    merged_df.columns = merged_df.columns.str.capitalize()
    merged_df.columns = merged_df.columns.str.replace("_", " ")

    return merged_df