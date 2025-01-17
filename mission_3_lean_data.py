import pandas as pd
from matplotlib import pyplot as plt
from datetime import datetime
from collections import defaultdict

# Constants
dev = 'Development'
done = 'Done'
artifacts = ['D&C', 'HLTP', 'HLTC', 'LLTC', 'LLTP']
features = ['Approach', 'Airport', 'Arrival', 'Constraints', 'Departure']

# Plot Setup
plt.style.use('seaborn-whitegrid')
plt.rcParams["figure.autolayout"] = True

# File Path
path = r'C:\Users\223050503\Downloads\Mission 3.xlsx'

# Read Excel Data
df = pd.read_excel(path, header=0)

# Utility Functions
def calculate_cycle_times(task_name, bucket_name, label_name, artifact):
    done_dict, dev_dict = {}, {}
    for i, task in enumerate(task_name):
        if done in bucket_name[i] and artifact in label_name[i]:
            done_dict[task] = current_date[i]
        elif dev in bucket_name[i] and artifact in label_name[i]:
            dev_dict[task] = current_date[i]

    cycle_times = {}
    for task in done_dict:
        if task in dev_dict:
            done_date = datetime.strptime(done_dict[task], '%m/%d/%y')
            dev_date = datetime.strptime(dev_dict[task], '%m/%d/%y')
            cycle_times[done_dict[task]] = (done_date - dev_date).days

    return cycle_times

def plot_cycle_times(cycle_times, title, filename):
    x = sorted(cycle_times.keys())
    y = [cycle_times[date] for date in x]

    median = calculate_median(y)

    plt.plot(x, y, 'o')
    plt.axhline(median, label='Median Cycle Time', linestyle='--')
    plt.legend(loc='upper right')
    plt.title(title)
    plt.xlabel('Completion Date')
    plt.xticks(rotation=90, fontsize='medium')
    plt.ylabel('Cycle Time Days')
    plt.savefig(filename)
    plt.clf()

    print(f"{title} median cycle time: {median} days")

def calculate_median(values):
    values.sort()
    n = len(values)
    if n % 2 == 0:
        return (values[n // 2] + values[n // 2 - 1]) // 2
    else:
        return values[n // 2]

def calculate_delivery_stats(task_name, bucket_name, label_name, due_date, artifact):
    due_done_dict = {}
    for i, task in enumerate(task_name):
        if done in bucket_name[i] and artifact in label_name[i]:
            due_done_dict[task] = (current_date[i], due_date[i])

    delta_times = {}
    ontime_list, late_list = [], []

    for task, (done_date, due_date) in due_done_dict.items():
        done_date = datetime.strptime(done_date, '%m/%d/%y')
        due_date = datetime.strptime(due_date, '%m/%d/%y')
        delta_time = (done_date - due_date).days
        delta_times[done_date.strftime('%m/%d/%y')] = delta_time

        if delta_time <= 3:
            ontime_list.append(delta_time)
        else:
            late_list.append(delta_time)

    num_due = len(delta_times)
    percent_ontime = round(len(ontime_list) / num_due * 100, 1) if num_due else 0
    percent_late = round(len(late_list) / num_due * 100, 1) if num_due else 0

    return delta_times, percent_ontime, percent_late

# Extract Data Columns
task_name = df['Task Name']
bucket_name = df['Bucket Name']
label_name = df['Labels']
due_date = pd.to_datetime(df['Due Date']).dt.strftime('%m/%d/%y')
current_date = pd.to_datetime(df['Current Date']).dt.strftime('%m/%d/%y')

# Process Each Artifact
for artifact in artifacts:
    cycle_times = calculate_cycle_times(task_name, bucket_name, label_name, artifact)
    plot_cycle_times(cycle_times, f"Cycle Time of {artifact} Items", f"mission3_{artifact}_cycle.png")

for feature in features:
    delivery_stats, percent_ontime, percent_late = calculate_delivery_stats(
        task_name, bucket_name, label_name, due_date, feature
    )
    plot_cycle_times(delivery_stats, f"{feature} On Time Delivery", f"mission3_{feature}_delivery.png")
    print(f"{feature} - On time: {percent_ontime}%, Late: {percent_late}%"
