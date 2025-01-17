# Cycle Time Analysis and Visualization

This project processes Microsoft Teams data to calculate cycle times, analyze delivery performance, and create visualizations. The Excel data used in this project is sourced from a General Electric (GE) internship, where data from Microsoft Teams was leveraged for analysis.

---

## Contents

### 1. **`barchart_mission3.py`**
   - **Purpose**: Visualizes weekly delivery performance with on-time and late task counts.
   - **Key Features**:
     - Reads data from an Excel file.
     - Processes task data to calculate on-time and late deliveries.
     - Generates a bar chart showing weekly delivery statistics.
   - **Output**:
     - A bar chart saved as `overall_late_bar.png`.

### 2. **`excel_to_dic.py`**
   - **Purpose**: Calculates cycle times for completed and Work-In-Progress (WIP) tasks and visualizes results.
   - **Key Features**:
     - Cleans task names by removing emojis.
     - Calculates cycle times for "Done" and WIP tasks.
     - Generates a plot for completed task cycle times.
   - **Output**:
     - A console log of average cycle times.
     - A plot showing completed task cycle times.

### 3. **`lean_data_from_teams.py`**
   - **Purpose**: Command-line utility for processing and analyzing Excel data.
   - **Key Features**:
     - Accepts an Excel file path as input.
     - Cleans task names and calculates cycle times.
     - Provides average cycle times and detailed task breakdowns.
   - **Usage**:
     - Run with: `python lean_data_from_teams.py <path_to_excel_file>`.

### 4. **`mission_3_lean_data.py`**
   - **Purpose**: Analyzes and visualizes cycle times and delivery statistics for specific artifacts and features.
   - **Key Features**:
     - Processes data to calculate cycle times for artifacts and features.
     - Computes on-time and late delivery percentages.
     - Creates plots for each category.
   - **Output**:
     - Plots showing cycle times and on-time delivery performance.
     - Console output of delivery statistics.

---

## Prerequisites

- Python 3.7+
- Required libraries:
  - `pandas`
  - `matplotlib`
  - `openpyxl`

Install dependencies with:
```bash
pip install pandas matplotlib openpyxl
```

---

## Usage Instructions

1. Ensure the required Excel files are present in the specified paths.
   - Update the `path` variable in the scripts to point to your Excel files.
2. Run the desired script using Python:
   ```bash
   python script_name.py
   ```
3. Outputs will be generated as visual plots or console logs.

---

## File Structure

- **Input Data**:
  - Excel files containing task details (e.g., `Mission 2.xlsx`, `Mission 3.xlsx`). These files were generated during a GE internship, leveraging data extracted from Microsoft Teams.
    
- **Scripts**:
  - Python scripts for specific analyses.
- **Outputs**:
  - Plots and console logs summarizing results.
