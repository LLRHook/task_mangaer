import datetime
import time
import pandas as pd
import win32gui

def get_active_window_title():
    window = win32gui.GetForegroundWindow()
    return win32gui.GetWindowText(window)

def track_app_usage(min_duration=1):
    current_app_name = None
    current_app_start_time = None
    usage_records = []

    try:
        while True:
            new_app_name = get_active_window_title()
            if new_app_name != current_app_name:
                if current_app_name is not None:
                    usage_end_time = datetime.datetime.now()
                    duration_seconds = (usage_end_time - current_app_start_time).total_seconds()
                    if duration_seconds >= min_duration:
                        usage_records.append({
                            'Application': current_app_name,
                            'Start Time': current_app_start_time,
                            'End Time': usage_end_time,
                            'Duration': convert_seconds_to_hms(duration_seconds)
                        })

                current_app_name = new_app_name
                current_app_start_time = datetime.datetime.now()

            time.sleep(1)  # Check every 1 second for better accuracy

    except KeyboardInterrupt:
        print("Stopping application tracking...")

    return usage_records

def convert_seconds_to_hms(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f'{int(hours)}:{int(minutes):02d}:{int(seconds):02d}'

def save_to_excel(data):
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
    filename = f'{timestamp}_app_usage.xlsx'

    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)

def main():
    usage_data = track_app_usage()
    save_to_excel(usage_data)

if __name__ == "__main__":
    main()
