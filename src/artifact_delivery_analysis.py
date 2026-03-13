import argparse
import os
import sys
import pandas as pd
from matplotlib import pyplot as plt
from datetime import datetime
from collections import defaultdict

# Plot setup
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams["figure.autolayout"] = True

# Constants
dev = 'Development'
done = 'Done'
artifacts = ['D&C', 'HLTP', 'HLTC', 'LLTC', 'LLTP']
features = ['Approach', 'Airport', 'Arrival', 'Constraints', 'Departure']


def get_user_inputs():
    parser = argparse.ArgumentParser(description="Cycle time and delivery analysis from Teams Excel export")
    parser.add_argument('excelfile', help='Path to the Mission 3 Excel file')
    args = parser.parse_args()
    if not os.path.isfile(args.excelfile):
        print(f'ERROR: {args.excelfile} not found')
        sys.exit(1)
    return args.excelfile


def calculate_median(values):
    values = sorted(values)
    n = len(values)
    if n % 2 == 0:
        return (values[n // 2] + values[n // 2 - 1]) // 2
    else:
        return values[n // 2]


def calculate_cycle_times(task_name, bucket_name, label_name, current_date, artifact=None):
    """3-pass dedup logic from original — preserves correct task matching."""
    done_dict, dev_dict = {}, {}
    current_task = ''

    # Pass 1: find done tasks (optionally filtered by artifact label)
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            match = done in bucket_name[i] and task_name[i] == task_name[j]
            if artifact:
                match = match and artifact in str(label_name[i])
            if match:
                if current_task != task_name[i]:
                    done_dict[task_name[i]] = current_date[i]
                    current_task = task_name[i]
                break

    current_task = ''

    # Pass 2: find dev tasks for tasks already in done_dict
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            match = dev in bucket_name[i] and task_name[i] == task_name[j] and task_name[i] in done_dict
            if artifact:
                match = match and artifact in str(label_name[i])
            if match:
                if current_task != task_name[i]:
                    dev_dict[task_name[i]] = current_date[i]
                    current_task = task_name[i]
                break

    current_task = ''
    done_dict = {}

    # Pass 3: rebuild done_dict filtered to tasks that have both done and dev entries
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            match = done in bucket_name[i] and task_name[i] == task_name[j] and task_name[i] in dev_dict
            if artifact:
                match = match and artifact in str(label_name[i])
            if match:
                if current_task != task_name[i]:
                    done_dict[task_name[i]] = current_date[i]
                    current_task = task_name[i]
                break

    # Calculate cycle times
    x, y = [], []
    for task in done_dict:
        c = datetime.strptime(done_dict[task], '%m/%d/%y')
        d = datetime.strptime(dev_dict[task], '%m/%d/%y')
        cycle_time = (c - d).days
        x.append(done_dict[task])
        x = sorted(x)
        y.append(cycle_time)

    return x, y


def plot_cycle_times(x, y, title, filename):
    median = calculate_median(y)
    ax = plt.subplot()
    plt.plot(x, y, 'o')
    ax.axhline(median, label='Median Cycle Time', linestyle='--')
    plt.legend(loc='upper right')
    plt.title(title)
    plt.xlabel('Completion Date')
    plt.xticks(rotation=90, fontsize='medium')
    plt.ylabel('Cycle Time Days')
    plt.savefig(filename)
    plt.clf()
    print(f"{title} median cycle time: {median} days")


def calculate_delivery_stats(task_name, bucket_name, label_name, due_date, current_date, feature):
    """3-pass dedup logic from original for delivery stats."""
    due_done_dict, due_date_dict = {}, {}
    current_task = ''

    # Pass 1: find done tasks for this feature
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            if done in bucket_name[i] and feature in str(label_name[i]) and task_name[i] == task_name[j]:
                if current_task != task_name[i]:
                    due_done_dict[task_name[i]] = current_date[i]
                    current_task = task_name[i]
                break

    # Pass 2: find due dates for those tasks
    current_task = ''
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            if done in bucket_name[i] and feature in str(label_name[i]) and task_name[i] == task_name[j] and task_name[i] in due_done_dict:
                if current_task != task_name[i]:
                    due_date_dict[task_name[i]] = due_date[i]
                    current_task = task_name[i]
                break

    current_task = ''
    due_done_dict = {}

    # Pass 3: rebuild filtered to tasks with both done and due dates
    for j in range(len(task_name)):
        for i in range(len(task_name)):
            if done in bucket_name[i] and feature in str(label_name[i]) and task_name[i] == task_name[j] and task_name[i] in due_date_dict:
                if current_task != task_name[i]:
                    due_done_dict[task_name[i]] = current_date[i]
                    current_task = task_name[i]
                break

    x, y = [], []
    ontime_list, late_list = [], []

    for task in due_done_dict:
        c = datetime.strptime(due_done_dict[task], '%m/%d/%y')
        d = datetime.strptime(due_date_dict[task], '%m/%d/%y')
        delta_time = (c - d).days
        if delta_time <= 3:
            ontime_list.append(delta_time)
        else:
            late_list.append(delta_time)
        x.append(due_done_dict[task])
        x = sorted(x)
        y.append(delta_time)

    num_due = len(due_done_dict)
    percent_ontime = round(len(ontime_list) / num_due * 100, 1) if num_due else 0
    percent_late = round(len(late_list) / num_due * 100, 1) if num_due else 0

    return x, y, len(ontime_list), len(late_list), num_due, percent_ontime, percent_late


def main():
    path = get_user_inputs()

    df = pd.read_excel(path, header=0)
    df2 = df.dropna(subset=['Labels'])
    df4 = df.dropna(subset=['Due Date', 'Labels'])

    # --- Overall cycle time (no artifact filter) ---
    task_name = list(df2.iloc[:, 0])
    bucket_name = list(df2.iloc[:, 2])
    label_name = list(df2.iloc[:, 16])
    current_date = pd.to_datetime(list(df2.iloc[:, 17])).strftime('%m/%d/%y')

    x, y = calculate_cycle_times(task_name, bucket_name, label_name, current_date, artifact=None)
    plot_cycle_times(x, y, 'Cycle Time of Done Items', 'mission3_overall_cycle.png')

    # --- Cycle time per artifact ---
    for artifact in artifacts:
        task_name = list(df2.iloc[:, 0])
        bucket_name = list(df2.iloc[:, 2])
        label_name = list(df2.iloc[:, 16])
        current_date = pd.to_datetime(list(df2.iloc[:, 17])).strftime('%m/%d/%y')

        x, y = calculate_cycle_times(task_name, bucket_name, label_name, current_date, artifact=artifact)
        plot_cycle_times(x, y, f'Cycle Time of {artifact} Items', f'mission3_{artifact}_cycle.png')

    # --- Delivery stats per feature ---
    for feature in features:
        task_name = list(df4.iloc[:, 0])
        bucket_name = list(df4.iloc[:, 2])
        label_name = list(df4.iloc[:, 16])
        due_date = pd.to_datetime(list(df4.iloc[:, 9])).strftime('%m/%d/%y')
        current_date = pd.to_datetime(list(df4.iloc[:, 17])).strftime('%m/%d/%y')

        x, y, num_ontime, num_late, num_due, pct_ontime, pct_late = calculate_delivery_stats(
            task_name, bucket_name, label_name, due_date, current_date, feature
        )
        plot_cycle_times(x, y, f'{feature} On Time Delivery', f'mission3_{feature}_delivery.png')

        print(f"Our {feature} arrival time median is: {calculate_median(y)} days")
        print(f"Number of {feature} early or on time items: {num_ontime} items")
        print(f"Number of {feature} late items: {num_late} items")
        print(f"Percentage of {feature} early or on time items is: {pct_ontime}%")
        print(f"Percentage of {feature} late items is: {pct_late}%")


if __name__ == '__main__':
    main()
