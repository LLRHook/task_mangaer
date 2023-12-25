import os
import datetime
import calendar
import time
import pandas as pd
import win32gui

def get_active_window_title():
    window = win32gui.GetForegroundWindow()
    return win32gui.GetWindowText(window)

def convert_seconds_to_hms(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f'{int(hours)}:{int(minutes):02d}:{int(seconds):02d}'

def create_directory(year, month):
    month_name = calendar.month_name[month]
    directory = os.path.join(year, month_name)
    if not os.path.exists(directory):
        os.makedirs(directory)
    return directory

def get_file_path():
    now = datetime.datetime.now()
    year = str(now.year)
    month = now.month
    day = now.day
    directory = create_directory(year, month)
    # Updated file naming convention
    filename = f'{day}_{now.strftime("%H-%M-%S")}_usage.xlsx'
    return os.path.join(directory, filename)

def track_app_usage():
    usage_records = []
    current_app_name = None
    current_app_start_time = None
    current_date = datetime.date.today()

    try:
        while True:
            new_app_name = get_active_window_title()
            new_date = datetime.date.today()
            
            if current_app_name and (new_app_name != current_app_name or new_date != current_date):
                usage_end_time = datetime.datetime.now()
                duration_seconds = (usage_end_time - current_app_start_time).total_seconds()
                
                # Format the start and end times to only include hours, minutes, and seconds
                formatted_start_time = current_app_start_time.strftime("%H:%M:%S")
                formatted_end_time = usage_end_time.strftime("%H:%M:%S")

                usage_records.append({
                    'Application': current_app_name,
                    'Start Time': formatted_start_time,
                    'End Time': formatted_end_time,
                    'Duration': convert_seconds_to_hms(duration_seconds)
                })

                if new_date != current_date:
                    save_to_excel(usage_records, get_file_path())
                    usage_records = []
                    current_date = new_date

                current_app_name = new_app_name
                current_app_start_time = datetime.datetime.now()

            elif not current_app_name:
                current_app_name = new_app_name
                current_app_start_time = datetime.datetime.now()

            time.sleep(1)  # Check every 1 second

    except KeyboardInterrupt:
        print("Stopping application tracking...")
        if usage_records:
            save_to_excel(usage_records, get_file_path())


def save_to_excel(data, file_path):
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)

def main():
    track_app_usage()

if __name__ == "__main__":
    main()
