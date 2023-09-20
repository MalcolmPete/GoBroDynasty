import pandas as pd

# Load the Excel file into a Pandas DataFrame
df = pd.read_excel('player_info.xlsx')

# Fill empty player values with zeroes
df['player_value'].fillna(0, inplace=True)

# Group players by username
grouped = df.groupby('username')

# Create an Excel writer to save data to an Excel file
output_file = 'teams.xlsx'

# Create an Excel writer with XlsxWriter engine
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    workbook = writer.book  # Get the workbook

    # Iterate through each group (team) and create sheets
    for username, group in grouped:
        # Sort players within the team by player value
        group = group.sort_values(by='player_value', ascending=False)

        # Create a new Excel sheet for each team
        group.to_excel(writer, sheet_name=username, index=False)

    # Create a "Summary" sheet
    summary_sheet = workbook.add_worksheet('Summary')

    # Initialize row and column counters for the "Summary" sheet
    summary_row = 0
    summary_col = 0

    # Write headers for the "Summary" sheet
    summary_sheet.write(summary_row, summary_col, 'Username')
    summary_sheet.write(summary_row, summary_col + 1, 'TeamValue')
    summary_row += 1

    # Create a list to store summary data for sorting
    summary_data = []

    # Iterate through each group (team) to calculate and write the total values
    for username, group in grouped:
        total_value = group['player_value'].sum()
        summary_data.append((username, total_value))  # Store data for sorting

    # Sort the summary data by team value in descending order
    summary_data.sort(key=lambda x: x[1], reverse=True)

    # Write sorted summary data to the "Summary" sheet
    for username, total_value in summary_data:
        summary_sheet.write(summary_row, summary_col, username)
        summary_sheet.write(summary_row, summary_col + 1, total_value)
        summary_row += 1

print(f"Teams and Summary have been exported to {output_file}")