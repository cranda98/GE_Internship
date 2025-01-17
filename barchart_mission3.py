import pandas as pd
from matplotlib import pyplot as plt
from datetime import datetime, timedelta
from collections import defaultdict, OrderedDict
import numpy as np

# Constants
DEV = 'Development'
DONE = 'Done'

# Artifacts
DC = 'D&C'
HLTP = 'HLTP'
LLTC = 'LLTC'
LLTP = 'LLTP'

# Features
APPROACH = 'Approach Vref Entry'
AIRPORT = 'Airport entries'
ARRIVAL = 'Arrival entries'
CONSTRAINTS = 'Constraints'
DEPARTURE = 'Departure entries'

# Plot setup
plt.style.use('seaborn-whitegrid')
plt.rcParams["figure.autolayout"] = True

# Path to Excel file
path = r'C:\Users\223050503\Downloads\Mission 3.xlsx'

# Variables
done_dict = {}
dev_dict = {}
due_done_dict = {}
due_date_dict = {}
ontime_tasks_by_week = defaultdict(int)
late_tasks_by_week = defaultdict(int)

# Load data from Excel
df = pd.read_excel(path, index_col=0, header=0)
df2 = df.dropna(subset=['Due Date'])

# Extract data
task_name = list(df2.iloc[:, 0])
bucket_name = list(df2.iloc[:, 1])
due_date = pd.to_datetime(list(df2.iloc[:, 8])).strftime('%m/%d/%y')
current_date = pd.to_datetime(list(df2.iloc[:, 16])).strftime('%m/%d/%y')

# Process tasks
current_task = ''
for j in range(len(task_name)):
    found = False
    for i in range(len(task_name)):
        if DONE in bucket_name[i] and task_name[i] == task_name[j]:
            found = True
            break
    if found and current_task != task_name[i]:
        due_done_dict[task_name[i]] = current_date[i]
        current_task = task_name[i]

# Map due dates
for j in range(len(task_name)):
    found = False
    for i in range(len(task_name)):
        if DONE in bucket_name[i] and task_name[i] == task_name[j] and task_name[i] in due_done_dict:
            found = True
            break
    if found and current_task != task_name[i]:
        due_date_dict[task_name[i]] = due_date[i]
        current_task = task_name[i]

# Update due_done_dict
current_task = ''
due_done_dict = {}
for j in range(len(task_name)):
    found = False
    for i in range(len(task_name)):
        if DONE in bucket_name[i] and task_name[i] == task_name[j] and task_name[i] in due_date_dict:
            found = True
            break
    if found and current_task != task_name[i]:
        due_done_dict[task_name[i]] = current_date[i]
        current_task = task_name[i]

# Calculate task timing
for task in due_done_dict:
    actual_date = datetime.strptime(due_done_dict[task], '%m/%d/%y')
    expected_date = datetime.strptime(due_date_dict[task], '%m/%d/%y')
    delta_time = (actual_date - expected_date).days
    week_num = (actual_date - timedelta(days=actual_date.weekday())).strftime('%m/%d/%y')

    if delta_time <= 3:
        ontime_tasks_by_week[week_num] += 1
    else:
        late_tasks_by_week[week_num] += 1

# Sort and organize data
ontime_tasks_by_week = dict(OrderedDict(sorted(ontime_tasks_by_week.items())))
late_tasks_by_week = dict(OrderedDict(sorted(late_tasks_by_week.items())))

weeks = sorted(ontime_tasks_by_week.keys())
ontime_items = list(ontime_tasks_by_week.values())
late_items = list(late_tasks_by_week.values())

# Bar chart configuration
bar_width = 0.35
positions = np.arange(len(weeks))
legend_labels = ['Ontime', 'Late']

plt.bar(positions, ontime_items, bar_width, label='Ontime', color='green', edgecolor='black')
plt.bar(positions + bar_width, late_items, bar_width, label='Late', color='red', edgecolor='black')

plt.legend(loc='upper left')
plt.title('Weekly Delivery Delta')
plt.xlabel('Completion Week')
plt.xticks(positions + bar_width / 2, weeks, rotation=90, fontsize='medium')
plt.ylabel('Items Completed')
plt.tight_layout()

# Save and show chart
plt.savefig('overall_late_bar.png')
plt.show()
