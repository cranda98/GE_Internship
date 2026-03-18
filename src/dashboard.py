"""
dashboard.py — GE Internship Lean Analytics Dashboard
Generates a two-panel visualization:
  - Left:  Cycle time distribution by artifact type (box plot)
  - Right: On-time delivery rate by feature area (horizontal bar chart)

Usage:
    python src/dashboard.py data/Mission3_rich_sample.xlsx
"""

import argparse
import os
import sys
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# ── Constants ────────────────────────────────────────────────────────────────
ARTIFACTS = ['D&C', 'HLTP', 'HLTC', 'LLTC', 'LLTP']
FEATURES  = ['Approach', 'Airport', 'Arrival', 'Constraints', 'Departure']
DEV_BUCKET  = 'Development'
DONE_BUCKET = 'Done'
ONTIME_THRESHOLD_DAYS = 3   # tasks completed within 3 days of due date = on time
COLORS = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']


# ── Argument parsing ─────────────────────────────────────────────────────────
def get_args():
    parser = argparse.ArgumentParser(description="Lean analytics dashboard from Teams Planner export")
    parser.add_argument('excelfile', help='Path to the Mission 3 Excel file')
    parser.add_argument('--output', default='ge_analytics_dashboard.png',
                        help='Output PNG filename (default: ge_analytics_dashboard.png)')
    args = parser.parse_args()
    if not os.path.isfile(args.excelfile):
        print(f'ERROR: {args.excelfile} not found')
        sys.exit(1)
    return args


# ── Data loading ─────────────────────────────────────────────────────────────
def load_data(path):
    df = pd.read_excel(path, header=0)
    df.columns = [
        'Task Name', 'C1', 'Bucket Name', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8',
        'Due Date', 'C10', 'C11', 'C12', 'C13', 'C14', 'C15', 'Labels', 'Current Date'
    ]
    df['Current Date'] = pd.to_datetime(df['Current Date'], infer_datetime_format=True)
    df['Due Date']     = pd.to_datetime(df['Due Date'],     infer_datetime_format=True)
    return df


# ── Cycle time per artifact ───────────────────────────────────────────────────
def cycle_times_for_artifact(df, artifact):
    """
    For each task that has both a Development row and a Done row (filtered by
    artifact label), return the cycle time in days.
    """
    mask = df['Labels'].astype(str).str.contains(artifact, na=False)
    sub  = df[mask].copy()

    dev_rows  = sub[sub['Bucket Name'].str.contains(DEV_BUCKET,  na=False)]
    done_rows = sub[sub['Bucket Name'].str.contains(DONE_BUCKET, na=False)]

    dev_dates  = dev_rows.groupby('Task Name')['Current Date'].min()
    done_dates = done_rows.groupby('Task Name')['Current Date'].max()

    common = dev_dates.index.intersection(done_dates.index)
    cycle  = (done_dates[common] - dev_dates[common]).dt.days
    return cycle[cycle >= 0].values


# ── On-time delivery per feature ──────────────────────────────────────────────
def delivery_stats_for_feature(df, feature):
    """
    For tasks in the Done bucket for a given feature, compare completion date
    to due date. Returns (pct_ontime, pct_late).
    """
    mask = (
        df['Bucket Name'].str.contains(DONE_BUCKET, na=False) &
        df['Labels'].astype(str).str.contains(feature, na=False) &
        df['Due Date'].notna()
    )
    sub = df[mask].drop_duplicates('Task Name')

    if sub.empty:
        return 0.0, 100.0

    delta = (sub['Current Date'] - sub['Due Date']).dt.days
    ontime = (delta <= ONTIME_THRESHOLD_DAYS).sum()
    total  = len(delta)
    pct_ontime = round(ontime / total * 100, 1)
    return pct_ontime, round(100 - pct_ontime, 1)


# ── Plotting ──────────────────────────────────────────────────────────────────
def make_dashboard(df, output_path):
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle('GE Internship — Lean Delivery Analytics', fontsize=15, fontweight='bold')

    # ── Left panel: cycle time box plot ──────────────────────────────────────
    ax1 = axes[0]
    data   = [cycle_times_for_artifact(df, art) for art in ARTIFACTS]
    medians = [int(np.median(d)) if len(d) else 0 for d in data]

    bp = ax1.boxplot(
        data,
        tick_labels=ARTIFACTS,
        patch_artist=True,
        notch=False,
        medianprops=dict(color='black', linewidth=2),
        flierprops=dict(marker='o', markersize=4, alpha=0.5),
    )
    for patch, color in zip(bp['boxes'], COLORS):
        patch.set_facecolor(color)
        patch.set_alpha(0.7)

    overall_median = int(np.median(np.concatenate([d for d in data if len(d)])))
    ax1.axhline(overall_median, color='gray', linestyle='--', linewidth=1.2,
                alpha=0.7, label=f'Team Median ({overall_median} days)')
    ax1.set_title('Cycle Time by Artifact Type', fontsize=12, fontweight='bold')
    ax1.set_xlabel('Artifact Type')
    ax1.set_ylabel('Cycle Time (Days)')
    ax1.set_ylim(bottom=0)
    ax1.legend(fontsize=9)
    ax1.grid(axis='y', alpha=0.3)

    # ── Right panel: on-time delivery horizontal bar ──────────────────────────
    ax2 = axes[1]
    stats   = [delivery_stats_for_feature(df, f) for f in FEATURES]
    ontime  = [s[0] for s in stats]
    late    = [s[1] for s in stats]
    y       = np.arange(len(FEATURES))

    ax2.barh(y, ontime, color='#2ca02c', alpha=0.85, label='On Time')
    ax2.barh(y, late,   left=ontime, color='#d62728', alpha=0.85, label='Late')

    for i, (ot, lt) in enumerate(zip(ontime, late)):
        if ot > 8:
            ax2.text(ot / 2,      i, f'{ot}%', va='center', ha='center',
                     fontsize=9, fontweight='bold', color='white')
        if lt > 8:
            ax2.text(ot + lt / 2, i, f'{lt}%', va='center', ha='center',
                     fontsize=9, fontweight='bold', color='white')

    ax2.set_yticks(y)
    ax2.set_yticklabels(FEATURES)
    ax2.set_xlabel('Percentage of Tasks (%)')
    ax2.set_title('On-Time Delivery Rate by Feature Area', fontsize=12, fontweight='bold')
    ax2.set_xlim(0, 100)
    ax2.axvline(70, color='gray', linestyle='--', linewidth=1.2,
                alpha=0.7, label='Target (70%)')
    ax2.legend(fontsize=9)
    ax2.grid(axis='x', alpha=0.3)

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    print(f'✓ Dashboard saved to {output_path}')


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    args = get_args()
    df   = load_data(args.excelfile)
    make_dashboard(df, args.output)


if __name__ == '__main__':
    main()
