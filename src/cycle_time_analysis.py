import argparse
import os
import sys
import pandas as pd
from datetime import datetime
from matplotlib import pyplot as plt
import re

# Plot setup
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams["figure.autolayout"] = True

emoji_pattern = re.compile("["
                           u"\U0001F600-\U0001F64F"  # emoticons
                           u"\U0001F300-\U0001F5FF"  # symbols & pictographs
                           u"\U0001F680-\U0001F6FF"  # transport & map symbols
                           u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
                           "]+", flags=re.UNICODE)


def get_user_inputs():
    parser = argparse.ArgumentParser(description="Cycle time analysis from Teams Excel export")
    parser.add_argument('excelfile', help='Path to the Mission 2 Excel file')
    args = parser.parse_args()
    if not os.path.isfile(args.excelfile):
        print(f'ERROR: {args.excelfile} not found')
        sys.exit(1)
    return args.excelfile


def main():
    now = datetime.now()
    path = get_user_inputs()

    # Pull Teams Excel Data
    df = pd.read_excel(path, index_col=0, header=0)
    df2 = df.fillna(value=0)

    task_name = list(df2.iloc[:, 0])
    clean_task_name = [emoji_pattern.sub(r'', t) for t in task_name]

    bucket_data = list(df2.iloc[:, 1])
    pull_created_date = list(df2.iloc[:, 6])
    pull_completed_date = list(df2.iloc[:, 10])

    done_avg_dic = {}
    wip_avg_dic = {}
    x = []
    y = []

    # Loop to organize done data
    for i in range(len(task_name)):
        if pull_completed_date[i] == 0:
            pull_completed_date[i] = pull_created_date[i]
        a = datetime.strptime(pull_created_date[i], '%m/%d/%Y')
        b = datetime.strptime(pull_completed_date[i], '%m/%d/%Y')
        cycle_time = (b - a).days
        if cycle_time > 0:
            if bucket_data[i] == 'Done':
                x.append(b)
                y.append(cycle_time)
                done_avg_dic[i] = cycle_time

    # Find done avg
    done_avg = sum(done_avg_dic.values()) // len(done_avg_dic)

    # Plot done tasks
    ax = plt.subplot()
    plt.plot(x, y, 'o')
    ax.axhline(done_avg, label='Average Cycle Time of Completed items', linestyle='--')
    plt.legend(loc='upper right')
    plt.title('Cycle Time Chart of Completed items')
    plt.xlabel('Completion Date')
    plt.ylabel('Cycle Time (Days)')
    plt.show()

    print(f"Done cycle time average: {done_avg} days")

    # For all WIP items
    for i in range(len(task_name)):
        a = datetime.strptime(pull_created_date[i], '%m/%d/%Y')
        wip = (now - a).days
        if bucket_data[i] == 'In Work' or bucket_data[i] == 'Review':
            wip_avg_dic[i] = wip

    # Find WIP average cycle time
    wip_avg = sum(wip_avg_dic.values()) // len(wip_avg_dic)
    print(f"WIP cycle time average: {wip_avg} days")


if __name__ == '__main__':
    main()
