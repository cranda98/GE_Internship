# Lean Analytics: Cycle Time & Delivery Performance Toolset 📊

A suite of Python tools built during a **General Electric internship** to analyze software delivery performance from Microsoft Teams Planner data. The tools calculate cycle times, track on-time delivery rates, and generate visualizations used to support engineering team retrospectives.

---

## Overview

Engineering teams at GE used Microsoft Teams Planner to track task progress across sprints. This toolset was built to extract meaningful lean metrics from that data — specifically **cycle time** (how long tasks take from development to done) and **delivery delta** (how often tasks are completed on time vs. late). These metrics were used to identify bottlenecks and improve team throughput.

---

## Scripts

### `lean_data_from_teams.py`
General-purpose cycle time analyzer. Accepts any Teams Planner Excel export and calculates average cycle time across all tasks, with a scatter plot of completion dates vs. cycle time.

```bash
python lean_data_from_teams.py <path_to_excel_file>
```

**Output:** Scatter plot of cycle times + console summary of average cycle time per task.

---

### `excel_to_dic.py`
Calculates and visualizes cycle times separately for completed (`Done`) and in-progress (`In Work`, `Review`) tasks. Strips emoji from task names for clean processing.

```bash
python excel_to_dic.py <path_to_excel_file>
```

**Output:** Cycle time plot for completed tasks + WIP average cycle time in console.

---

### `mission_3_lean_data.py`
Deep-dive analysis broken down by artifact type (D&C, HLTP, HLTC, LLTC, LLTP) and feature area (Approach, Airport, Arrival, Constraints, Departure). Uses a 3-pass deduplication algorithm to correctly match tasks across bucket transitions.

```bash
python mission_3_lean_data.py <path_to_excel_file>
```

**Output:** Per-artifact cycle time charts + per-feature on-time delivery charts and percentages. Saves PNGs for each category.

---

### `barchart_mission3.py`
Generates a weekly bar chart of on-time vs. late task completions. Tasks completed within 3 days of their due date are counted as on-time.

```bash
python barchart_mission3.py <path_to_excel_file>
```

**Output:** `overall_late_bar.png` — weekly stacked bar chart of delivery performance.

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| Python 3.7+ | Core language |
| pandas | Excel ingestion and data manipulation |
| matplotlib | Visualization |
| openpyxl | Excel file reading |

---

## Setup

```bash
pip install pandas matplotlib openpyxl
```

---

## Input Data Format

All scripts expect a Microsoft Teams Planner Excel export with standard column ordering (Task Name, Bucket Name, Labels, Due Date, Created Date, Completed Date, Current Date). The Excel files are not included in this repo as they contain proprietary GE project data.

---

## Notes

Scripts were refactored after the internship for portability and clarity — hardcoded paths replaced with CLI arguments, logic wrapped in `main()`, and deprecated dependencies updated.
