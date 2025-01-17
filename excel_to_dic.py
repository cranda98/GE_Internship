# Imports
import pandas as pd
from datetime import datetime
from matplotlib import pyplot as plt
import re

# Current date and time
now = datetime.now()

# Plot setup
plt.style.use('seaborn-whitegrid')
plt.rcParams["figure.autolayout"] = True

# Path to the Excel file
path = r'C:\Users\223050503\Downloads\Mission 2.xlsx'

# Regex to strip emojis
emoji_pattern = re.compile(
    "["
    u"\U0001F600-\U0001F64F"  # Emoticons
    u"\U0001F300-\U0001F5FF"  # Symbols & pictographs
    u"\U0001F680-\U0001F6FF"  # Transport & map symbols
    u"\U0001F1E0-\U0001F1FF"  # Flags (iOS)
    "]+", flags=re.UNICODE
)

# Read and clean Excel data
df = pd.read_excel(path, index_col=0, header=0).fillna(value=0)

# Extract relevant columns
task_names = [emoji_pattern.sub('', name) for name in df.iloc[:, 0]]
bucket_data = list(df.iloc[:, 1])
created_dates = list(df.iloc[:, 6])
completed_dates = list(df.iloc[:, 10])
label_data = list(df.iloc[:, 15])

# Pre-define dictionaries and lists
done_avg_dic, wip_avg_dic = {}, {}
x, y = [], []

# Process "Done" tasks and calculate cycle times
for i, name in enumerate(task_names):
    if completed_dates[i] == 0:
        completed_dates[i] = created_dates[i]
    
    start_date = datetime.strptime(created_dates[i], '%m/%d/%Y')
    end_date = datetime.strptime(completed_dates[i], '%m/%d/%Y')
    cycle_time = (end_date - start_date).days

    if cycle_time > 0 and bucket_data[i] == 'Done':
        x.append(end_date)
        y.append(cycle_time)
        done_avg_dic[i] = cycle_time

# Calculate average cycle time for "Done" tasks
done_avg = sum(done_avg_dic.values()) // len(done_avg_dic)

# Plot "Done" tasks cycle times
plt.plot(x, y, 'o')
plt.axhline(done_avg, label='Average Cycle Time (Done)', linestyle='--')
plt.title('Cycle Time Chart of Completed Items')
plt.xlabel('Completion Date')
plt.ylabel('Cycle Time (Days)')
plt.legend(loc='upper right')
plt.show()

print(f"Done cycle time average: {done_avg} days")

# Process WIP tasks
for i, name in enumerate(task_names):
    start_date = datetime.strptime(created_dates[i], '%m/%d/%Y')
    wip_time = (now - start_date).days

    if bucket_data[i] in ['In Work', 'Review']:
        wip_avg_dic[i] = wip_time

# Calculate average cycle time for WIP tasks
wip_avg = sum(wip_avg_dic.values()) // len(wip_avg_dic)
print(f"WIP cycle time average: {wip_avg} days")
