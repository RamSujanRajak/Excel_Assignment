import pandas as pd

def analyze_excel_file(file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)  # Use the provided file path instead of a fixed path

    # Convert the 'Time' columns to datetime type for date calculations
    df['Time'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])
    df['Pay Cycle Start Date'] = pd.to_datetime(df['Pay Cycle Start Date'])
    df['Pay Cycle End Date'] = pd.to_datetime(df['Pay Cycle End Date'])

    # Sort the DataFrame by 'Employee Name', 'Position ID', and 'Time'
    df.sort_values(by=['Employee Name', 'Position ID', 'Time'], inplace=True)

    # Initialize variables for consecutive days and shift duration calculations
    consecutive_days_threshold = 7
    time_between_shifts_min = 60  # in minutes
    time_between_shifts_max = 600  # in minutes (10 hours)
    max_shift_duration = 840  # in minutes (14 hours)

    # Initialize lists to store Work
    less_than_10_hours_Work = []
    more_than_14_hours_Work = []
    consecutive_7_days_Work = []

    # Iterate over each employee
    for (employee_name, position_id), employee_data in df.groupby(['Employee Name', 'Position ID']):
        consecutive_days_count = 0
        previous_shift_end = None

        # Iterate over each row for the current employee
        for index, row in employee_data.iterrows():
            # Check for consecutive days
            if previous_shift_end is not None:
                days_between = (row['Time'] - previous_shift_end).days
                if days_between == 1:
                    consecutive_days_count += 1
                else:
                    consecutive_days_count = 0

            # Check for less than 10 hours between shifts
            if (
                previous_shift_end is not None
                and (row['Time'] - previous_shift_end).seconds / 60 < time_between_shifts_max 
                and (row['Time'] - previous_shift_end).seconds / 60 > time_between_shifts_min
            ):
                less_than_10_hours_Work.append(
                    f"{employee_name} ({position_id}): Less than 10 hours between shifts on {row['Time']}"
                )

            # Check for more than 14 hours in a single shift
            shift_duration = row['Time Out'] - row['Time']
            if shift_duration.seconds / 60 > max_shift_duration:
                more_than_14_hours_Work.append(
                    f"{employee_name} ({position_id}): More than 14 hours in a single shift on {row['Time']}"
                )

            # Update previous_shift_end for the next iteration
            previous_shift_end = row['Time Out']

            # Check if the employee has worked for 7 consecutive days
            if consecutive_days_count == consecutive_days_threshold:
                consecutive_7_days_Work.append(
                    f"{employee_name} ({position_id}): Worked for 7 consecutive days ending on {row['Time']}"
                )
                consecutive_days_count = 0  # Reset the count after appending

    # Print separate outputs for each type of Work
    if less_than_10_hours_Work:
        print("Less than 10 hours between shifts:")
        for Work in less_than_10_hours_Work:
            print(Work)

    if more_than_14_hours_Work:
        print("\nMore than 14 hours in a single shift:")
        for Work in more_than_14_hours_Work:
            print(Work)

    if consecutive_7_days_Work:
        print("\nWorked for 7 consecutive days:")
        for Work in consecutive_7_days_Work:
            print(Work)

# Take the Excel file path as input from the user
file_path = input("Enter the path of the Excel file: ")

# Call the analyze_excel_file function with the provided file path
analyze_excel_file(file_path)
