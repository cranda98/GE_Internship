import argparse
import os
import sys
import pandas as pd
from matplotlib import pyplot as plt
from datetime import datetime, timedelta
from collections import defaultdict, OrderedDict
import numpy as np

# Plot setup
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams["figure.autolayout"] = True


def get_user_inputs():
    parser = argparse.ArgumentParser(description="Weekly delivery delta bar chart from Teams Excel export")
    parser.add_argument('excelfile', help='Path to the Mission 3 Excel file')
    args = parser.parse_args()
    if not os.path.isfile(args.excelfile):
        print(f'ERROR: {args.excelfile} not found')
        sys.exit(1)
    return args.excelfile


def main():
    path = get_user_inputs()

    # Load data
    df = pd.read_excel(path, index_col=0, header=0)
    df2 = df.dropna(subset=['Due Date'])

    task_name = list(df2.iloc[:, 0])
    bucket_name = list(df2.iloc[:, 1])
    due_date = pd.to_datetime(list(df2.iloc[:, 8])).strftime('%m/%d/%y')
    current_date = pd.to_datetime(list(df2.iloc[:, 16])).strftime('%m/%d/%y')

    done = 'Done'
    due_done_dict = {}
    due_date_dict = {}
    ontime_tasks_by_week = defaultdict(int)
    late_tasks_by_week = defaultdict(int)

    # Pass 1: find completion dates for done tasks
    current_task = ''
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            if done in bucket_name[i] and task_name[i] == task_name[j]:
                if current_task != task_name[i]:
                    due_done_dict[task_name[i]] = current_date[i]
                    current_task = task_name[i]
                break

    # Pass 2: find due dates for those tasks
    current_task = ''
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            if done in bucket_name[i] and task_name[i] == task_name[j] and task_name[i] in due_done_dict:
                if current_task != task_name[i]:
                    due_date_dict[task_name[i]] = due_date[i]
                    current_task = task_name[i]
                break

    # Pass 3: rebuild due_done_dict filtered to tasks with due dates
    current_task = ''
    due_done_dict = {}
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            if done in bucket_name[i] and task_name[i] == task_name[j] and task_name[i] in due_date_dict:
                if current_task != task_name[i]:
                    due_done_dict[task_name[i]] = current_date[i]
                    current_task = task_name[i]
                break

    # Calculate on-time vs late by week
    for task, done_date_str in due_done_dict.items():
        done_dt = datetime.strptime(done_date_str, '%m/%d/%y')
        due_dt = datetime.strptime(due_date_dict[task], '%m/%d/%y')
        delta_days = (done_dt - due_dt).days
        week = (done_dt - timedelta(days=done_dt.weekday())).strftime('%m/%d/%y')

        if delta_days <= 3:
            ontime_tasks_by_week[week] += 1
        else:
            late_tasks_by_week[week] += 1

    # Sort and plot
    weeks = sorted(ontime_tasks_by_week.keys())
    ontime_items = [ontime_tasks_by_week[w] for w in weeks]
    late_items = [late_tasks_by_week[w] for w in weeks]

    pos = np.arange(len(weeks))
    bar_width = 0.35

    plt.bar(pos, ontime_items, bar_width, label='Ontime', color='green', edgecolor='black')
    plt.bar(pos + bar_width, late_items, bar_width, label='Late', color='red', edgecolor='black')

    plt.legend(loc='upper left')
    plt.title('Weekly Delivery Delta')
    plt.xlabel('Completion Week')
    plt.xticks(pos + bar_width / 2, weeks, rotation=90, fontsize='medium')
    plt.ylabel('Items Completed')

    plt.savefig('overall_late_bar.png')
    plt.show()


if __name__ == '__main__':
    main()
