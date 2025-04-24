import pandas as pd

capacity_sheets = pd.read_excel("Anon_Capacity Planning.xlsx", sheet_name=None)

def process_capacity_sheet(df, month_name):
    """Returns one capacity sheet"""

    #column for names
    df.columns = [str(col).strip() for col in df.columns]

    #employee and project columns
    id_cols = [col for col in df.columns  if col in ['Resource', 'Project']]

    #week columns
    week_cols = [col for col in df.columns if 'Week' in col]

    #wide -> long
    long_df = df.melt(id_vars = id_cols, value_vars = week_cols,
                      var_name = 'week', value_name = 'planned_days')

    #drop rows with no data
    long_df.dropna(subset=['planned_days'], inplace=True)

    # Convert commas to dots for decimal values (e.g. "0,5" → "0.5")
    long_df['planned_days'] = long_df['planned_days'].astype(str).str.replace(',', '.', regex=False)

    # Ensure planned_days is numeric
    long_df['planned_days'] = pd.to_numeric(long_df['planned_days'], errors='coerce')
    long_df.dropna(subset=['planned_days'], inplace=True)

    #month tag
    long_df['month'] = month_name

    #convert into hours
    long_df['planned_hours'] = long_df['planned_days'] * 8

    result = long_df.groupby(['Resource', 'Project', 'month'], as_index=False)['planned_hours'].sum()
    result.rename(columns={'Resource': 'employee', 'Project': 'project'}, inplace=True)

    return result

#nov_df = capacity_sheets['Resource Requirements - Nov']
#processed_nov = process_capacity_sheet(nov_df, 'Nov_2023')
#print(processed_nov)

all_months = []
#automatically process all monthly sheets
for sheet_name, df in capacity_sheets.items():
    if "Resource Requirements - " in sheet_name:
        #Extract month name
        month = sheet_name.split(" - ")[-1]
        month_tag = f"{month}"

        #process and store the result
        processed = process_capacity_sheet(df, month_tag)
        all_months.append(processed)

#combine all monthly data into a single DataFrame
planned_df = pd.concat(all_months, ignore_index = True)
#print(planned_df)




#Actual hours sheet
actual_df = pd.read_excel("Anon_Hours.xlsx", sheet_name = "Hours", skiprows = 2)
#TODO: get rid of hardcoded month names
actual_df.columns = ['employee', 'project', 'Nov_2023', 'Dec_2023', 'Jan_2024', 'Feb_2024', 'Total']
actual_df.dropna(subset=['employee', 'project'], inplace=True)
actual_df['employee'] = actual_df['employee'].astype(str)
actual_df['project'] = actual_df['project'].astype(str)

#melt into long format
actual_long = actual_df.melt(id_vars=['employee', 'project'],
                             value_vars=['Nov_2023', 'Dec_2023', 'Jan_2024', 'Feb_2024'],
                             var_name='month', value_name='actual_hours')
actual_long['actual_hours'] = actual_long['actual_hours'].astype(str).str.replace(',', '.', regex=False)
actual_long['actual_hours'] = pd.to_numeric(actual_long['actual_hours'], errors='coerce')
actual_long.dropna(subset=['actual_hours'], inplace=True)

print(actual_long)







