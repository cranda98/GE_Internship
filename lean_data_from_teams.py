# Imports
import argparse
import os
import pandas as pd
from datetime import datetime
from matplotlib import pyplot as plt
import re

# Constants
DEFAULT_EXCEL_INPUT_FILEPATH = 'Mission_2.xlsx'
now = datetime.now()

# Plot setup
plt.style.use('seaborn-whitegrid')

# Regex to strip emojis
emoji_pattern = re.compile(
    "["
    u"\U0001F600-\U0001F64F"
    u"\U0001F300-\U0001F5FF"
    u"\U0001F680-\U0001F6FF"
    u"\U0001F1E0-\U0001F1FF"
    "]+", flags=re.UNICODE
)

# Functions
def get_user_inputs():
    parser = argparse.ArgumentParser(description="Process Excel data for analysis")
    parser.add_argument('excelfile', default=DEFAULT_EXCEL_INPUT_FILEPATH, help='Path to the Excel file')
    args = parser.parse_args()

    if not os.path.isfile(args.excelfile):
        raise FileNotFoundError(f"File {args.excelfile} not found")
    return args.excelfile


def read_excel_file(filepath):
    return pd.read_excel(filepath, index_col=0, header=0).fillna(value=0)


def clean_task_names(df):
    return [emoji_pattern.sub('', name) for name in df.iloc[:, 0]]


def generate_cycle_time_data(task_names, create_dates, complete_dates, bucket_names):
    table_dic, avg_dic = {}, {}
    for i, name in enumerate(task_names):
        start_date = datetime.strptime(create_dates[i], '%m/%d/%Y')
        end_date = datetime.strptime(complete_dates[i] if complete_dates[i] != 0 else create_dates[i], '%m/%d/%Y')
        cycle_time = (end_date - start_date).days

        if cycle_time > 0:
            table_dic[name] = [bucket_names[i], cycle_time]
            avg_dic[name] = cycle_time

    return table_dic, avg_dic


def calculate_average(avg_dic):
    return sum(avg_dic.values()) // len(avg_dic)


# Main Function
def main():
    filepath = get_user_inputs()
    df = read_excel_file(filepath)

    task_names = clean_task_names(df)
    create_dates = df.iloc[:, 6].tolist()
    complete_dates = df.iloc[:, 10].tolist()
    bucket_names = df.iloc[:, 1].tolist()

    table_dic, avg_dic = generate_cycle_time_data(task_names, create_dates, complete_dates, bucket_names)
    avg_cycle_time = calculate_average(avg_dic)

    print(f"Average Cycle Time: {avg_cycle_time} days")
    for task, details in table_dic.items():
        print(f"{task}: {details}")


# Run Script
if __name__ == "__main__":
    main()
