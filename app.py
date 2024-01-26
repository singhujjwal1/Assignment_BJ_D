import pandas as pd
from datetime import datetime, timedelta

def parse_time(time_str):
    """
    Parse time from string to datetime object.
    """
    # Check if time_str is NaN (Not a Number)
    return datetime.strptime(str(time_str), "%Y-%m-%d %H:%M:%S") if not pd.isna(time_str) else None

def solve(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Print the column names to understand the structure
        output_text = "Columns: " + ', '.join(df.columns) + '\n\n'
        print("Columns:", df.columns)

        # Adjust the column names based on your actual file
        expected_columns = {'Position ID', 'Time', 'Time Out', 'Employee Name'}

        # Check if the required columns are present in the dataframe
        if not expected_columns.issubset(df.columns):
            output_text += "Error: The required columns are missing in the file."
            with open('output.txt', 'w') as output_file:
                output_file.write(output_text)
            return

        # Initialize sets to store unique results for each condition
        consecutive_days_result = set()
        less_than_10_hours_result = set()
        more_than_14_hours_result = set()

        # Dictionary to store employee data
        employees = {}

        # Iterate through each row in the dataframe
        for index, row in df.iterrows():
            try:
                # Extract relevant information from the row
                position_id = row['Position ID']
                employee_name = row['Employee Name']
                time_in = parse_time(row['Time'])
                time_out = parse_time(row['Time Out'])
            except KeyError as e:
                # Handle missing keys in the row
                continue

            # Check for invalid date/time values
            if time_in is None or time_out is None:
                continue

            # Create or update employee data in the dictionary
            if position_id not in employees:
                employees[position_id] = {'name': employee_name, 'shifts': []}
            employees[position_id]['shifts'].append((time_in, time_out))

        # Analyze shifts for each employee
        for position_id, data in employees.items():
            shifts = data['shifts']
            shifts.sort(key=lambda x: x[0])  # Sort shifts by start time

            for i in range(len(shifts) - 1):
                # Check for consecutive days
                if (shifts[i + 1][0] - shifts[i][1]).days == 1:
                    consecutive_days_result.add(f"{data['name']} ({position_id})")

                # Check for time between shifts
                time_between_shifts = shifts[i + 1][0] - shifts[i][1]
                if timedelta(hours=1) < time_between_shifts < timedelta(hours=10):
                    less_than_10_hours_result.add(f"{data['name']} ({position_id})")

            for shift in shifts:
                # Check for more than 14 hours in a single shift
                if (shift[1] - shift[0]).total_seconds() / 3600 > 14:
                    more_than_14_hours_result.add(f"{data['name']} ({position_id})")

        # Append the unique results for each condition to the output text
        output_text += "\nEmployees who worked for 7 consecutive days:\n"
        output_text += '\n'.join(consecutive_days_result) + '\n\n'

        output_text += "Employees with less than 10 hours between shifts:\n"
        output_text += '\n'.join(less_than_10_hours_result) + '\n\n'

        output_text += "Employees who worked for more than 14 hours in a single shift:\n"
        output_text += '\n'.join(more_than_14_hours_result) + '\n'


        # Print the unique results for each condition
        print("\nEmployees who worked for 7 consecutive days:")
        for result in consecutive_days_result:
            print(result)

        print("\nEmployees with less than 10 hours between shifts:")
        for result in less_than_10_hours_result:
            print(result)

        print("\nEmployees who worked for more than 14 hours in a single shift:")
        for result in more_than_14_hours_result:
            print(result)

        # Write the output text to the output file
        with open('output.txt', 'w') as output_file:
            output_file.write(output_text)

    except pd.errors.EmptyDataError:
        print("Error: The file is empty.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    file_path = "./sheet.xlsx"
    solve(file_path)
