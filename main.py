import json
import pandas as pd
from datetime import datetime, timezone
excel_file="Assignment_Timecard.xlsx"
df = pd.read_excel(excel_file)

json_data = df.to_json(orient='records')

json_data = json.loads(json_data)

results = {
    'consecutive_days_7': [],
    'less_than_10_hours_apart': [],
    'more_than_14_hours_single_shift': []
}


consecutive_days = 0
previous_end_time = None
current_employee = None


def epoch_to_utc(epoch_timestamp):
    try:
        utc_datetime = datetime.fromtimestamp(epoch_timestamp, timezone.utc)
        return utc_datetime
    except Exception as e:
        return None 
def is_within_range(start_date, end_date, days):
    if start_date is None or end_date is None:
        return False
    return (end_date - start_date) // (24 * 3600 * 1000) == days


def parse_time_to_minutes(time_str):
    if time_str is None:
        return 0  
    hours, minutes = map(int, time_str.split(':'))
    return hours * 60 + minutes

for data in json_data:
    employee_name = data['Employee Name']
    position = data['Position ID']
    time_in = data['Time']
    time_out = data['Time Out']
    timecard_hours = parse_time_to_minutes(data['Timecard Hours (as Time)'])
    if previous_end_time is not None and current_employee == employee_name:
        if is_within_range(previous_end_time, time_in, 1):
            consecutive_days += 1
        else:
            consecutive_days = 0
    else:
        consecutive_days = 0

    if consecutive_days == 7:
        results['consecutive_days_7'].append({'Employee Name': employee_name, 'Position ID': position, 'Time Out': time_out})

    if time_in is not None and current_employee == employee_name:
        if previous_end_time is not None and (time_in - previous_end_time) < (10 * 3600 * 1000) and (time_in - previous_end_time) > (1 * 3600 * 1000):
            results['less_than_10_hours_apart'].append({'Employee Name': employee_name, 'Position ID': position, 'Time Out': time_out})
    if timecard_hours > 14 * 60:
        results['more_than_14_hours_single_shift'].append({'Employee Name': employee_name, 'Position ID': position, 'Time Out': time_out})
    previous_end_time = time_out
    current_employee = employee_name


print("Employees who worked for 7 consecutive days:")
print(results['consecutive_days_7'])
print("\nEmployees with shifts less than 10 hours apart but greater than 1 hour:")
print(results['less_than_10_hours_apart'])
print("\nEmployees who worked more than 14 hours in a single shift:")
print(results['more_than_14_hours_single_shift'])

