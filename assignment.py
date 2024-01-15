import pandas as pd
from datetime import datetime, timedelta

def analyze_excel(input_file):
    # Read the Excel sheet into a pandas DataFrame
    df = pd.read_excel(input_file)

    # Dictionary to store employee information
    employees = {}

    # Loop through each row in the DataFrame
    for index, row in df.iterrows():
        name = row['Employee Name']
        date = row['Time']
        hours = row['Timecard Hours (as Time)']

        # Check consecutive days
        if name not in employees:
            employees[name] = {'consecutive_days': 0, 'last_date': None, 'last_hours': 0}

        # Convert date and hours to appropriate data types
        date = datetime.strptime(str(date), "%m/%d/%Y %I:%M %p")
        hours = sum(x * int(t) for x, t in zip([3600, 60], str(hours).split(':')))

        # Check consecutive days
        if employees[name]['last_date'] and (date - employees[name]['last_date']).days == 1:
            employees[name]['consecutive_days'] += 1
        else:
            employees[name]['consecutive_days'] = 1

        # Check less than 10 hours between shifts but greater than 1 hour
        if employees[name]['last_date'] and (date - employees[name]['last_date']).seconds < 10 * 3600 and (date - employees[name]['last_date']).seconds > 3600:
            print(f"{name} has less than 10 hours between shifts on {date}")

        # Check more than 14 hours in a single shift
        if hours > 14 * 3600:
            print(f"{name} has worked for more than 14 hours on {date}")

        # Update employee information
        employees[name]['last_date'] = date
        employees[name]['last_hours'] = hours

        # Check if employee has worked for 7 consecutive days
        if employees[name]['consecutive_days'] == 7:
            start_date = date - timedelta(days=6)
            print(f"{name} has worked for 7 consecutive days starting from {start_date} to {date}")

if __name__ == "__main__":
    # Assuming the input file is named 'employee_data.xlsx'
    input_file = 'input_file.xlsx'
    
    # Analyze the Excel sheet
    analyze_excel(input_file)
