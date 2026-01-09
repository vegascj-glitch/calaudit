"""
Chart generation for calendar audit dashboard.
Uses matplotlib for visualization.
"""

import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd
import numpy as np
from typing import Optional


# Color palette for consistent styling
COLORS = {
    'primary': '#4A90A4',
    'secondary': '#7CB9C7',
    'accent': '#E8A87C',
    'warning': '#C38D9E',
    'neutral': '#95A5A6',
    'background': '#F5F6FA',
}

# Weekday order for consistent display
WEEKDAY_ORDER = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']


def setup_style():
    """Configure matplotlib style for clean, professional charts."""
    plt.style.use('seaborn-v0_8-whitegrid')
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.size'] = 10
    plt.rcParams['axes.titlesize'] = 12
    plt.rcParams['axes.titleweight'] = 'bold'
    plt.rcParams['axes.labelsize'] = 10
    plt.rcParams['figure.facecolor'] = 'white'
    plt.rcParams['axes.facecolor'] = 'white'
    plt.rcParams['axes.edgecolor'] = '#CCCCCC'
    plt.rcParams['grid.color'] = '#EEEEEE'


def create_weekday_chart(df: pd.DataFrame) -> Optional[plt.Figure]:
    """
    Create a bar chart showing meeting load by weekday.

    Args:
        df: DataFrame with weekday distribution data

    Returns:
        matplotlib Figure object
    """
    setup_style()

    if df.empty:
        return None

    # Ensure proper weekday ordering
    df = df.copy()
    df['weekday'] = pd.Categorical(df['weekday'], categories=WEEKDAY_ORDER, ordered=True)
    df = df.sort_values('weekday')

    fig, ax = plt.subplots(figsize=(10, 5))

    bars = ax.bar(
        df['weekday'],
        df['total_hours'],
        color=COLORS['primary'],
        edgecolor='white',
        linewidth=1
    )

    # Add value labels on bars
    for bar, count in zip(bars, df['meeting_count']):
        height = bar.get_height()
        ax.annotate(
            f'{height:.1f}h\n({count} mtgs)',
            xy=(bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 3),
            textcoords="offset points",
            ha='center',
            va='bottom',
            fontsize=9
        )

    ax.set_xlabel('Day of Week')
    ax.set_ylabel('Total Hours')
    ax.set_title('Meeting Load by Day of Week')

    # Clean up axes
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    return fig


def create_duration_histogram(df: pd.DataFrame) -> Optional[plt.Figure]:
    """
    Create a histogram showing meeting duration distribution.

    Args:
        df: DataFrame with duration_minutes column

    Returns:
        matplotlib Figure object
    """
    setup_style()

    if df.empty or 'duration_minutes' not in df.columns:
        return None

    durations = df['duration_minutes'].dropna()
    if len(durations) == 0:
        return None

    fig, ax = plt.subplots(figsize=(10, 5))

    # Create bins at common meeting durations
    max_duration = min(durations.max(), 180)  # Cap at 3 hours for readability
    bins = [0, 15, 30, 45, 60, 90, 120, 180, max(180, durations.max() + 1)]

    n, bins_out, patches = ax.hist(
        durations.clip(upper=180),
        bins=bins,
        color=COLORS['primary'],
        edgecolor='white',
        linewidth=1
    )

    # Color long meetings differently
    for i, patch in enumerate(patches):
        if bins_out[i] >= 60:
            patch.set_facecolor(COLORS['accent'])

    ax.set_xlabel('Duration (minutes)')
    ax.set_ylabel('Number of Meetings')
    ax.set_title('Meeting Duration Distribution')

    # Add custom x-tick labels
    ax.set_xticks([15, 30, 45, 60, 90, 120, 180])
    ax.set_xticklabels(['15m', '30m', '45m', '1h', '1.5h', '2h', '3h+'])

    # Clean up axes
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # Add legend
    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor=COLORS['primary'], label='Standard (≤60 min)'),
        Patch(facecolor=COLORS['accent'], label='Long (>60 min)')
    ]
    ax.legend(handles=legend_elements, loc='upper right')

    plt.tight_layout()
    return fig


def create_daily_load_chart(df: pd.DataFrame) -> Optional[plt.Figure]:
    """
    Create a line chart showing meeting load over time.

    Args:
        df: DataFrame with daily load data

    Returns:
        matplotlib Figure object
    """
    setup_style()

    if df.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.fill_between(
        range(len(df)),
        df['total_hours'],
        alpha=0.3,
        color=COLORS['primary']
    )
    ax.plot(
        range(len(df)),
        df['total_hours'],
        color=COLORS['primary'],
        linewidth=2,
        marker='o',
        markersize=4
    )

    # Set x-axis labels (show every nth date for readability)
    n_labels = min(10, len(df))
    step = max(1, len(df) // n_labels)
    ax.set_xticks(range(0, len(df), step))
    ax.set_xticklabels(
        [str(d) for d in df['date'].iloc[::step]],
        rotation=45,
        ha='right'
    )

    ax.set_xlabel('Date')
    ax.set_ylabel('Hours in Meetings')
    ax.set_title('Daily Meeting Load Over Time')

    # Clean up axes
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    return fig


def create_top_meetings_chart(df: pd.DataFrame) -> Optional[plt.Figure]:
    """
    Create a horizontal bar chart of top meetings by time.

    Args:
        df: DataFrame with top meetings data

    Returns:
        matplotlib Figure object
    """
    setup_style()

    if df.empty:
        return None

    # Take top 8 for readability
    df = df.head(8).iloc[::-1]  # Reverse for horizontal bar chart

    fig, ax = plt.subplots(figsize=(10, 6))

    # Truncate long subjects
    subjects = df['subject'].apply(lambda x: x[:40] + '...' if len(x) > 40 else x)

    bars = ax.barh(
        subjects,
        df['total_hours'],
        color=COLORS['primary'],
        edgecolor='white',
        linewidth=1
    )

    # Add value labels
    for bar, occurrences in zip(bars, df['occurrences']):
        width = bar.get_width()
        ax.annotate(
            f'{width:.1f}h ({occurrences}x)',
            xy=(width, bar.get_y() + bar.get_height() / 2),
            xytext=(3, 0),
            textcoords="offset points",
            ha='left',
            va='center',
            fontsize=9
        )

    ax.set_xlabel('Total Hours')
    ax.set_title('Top Meetings by Total Time')

    # Clean up axes
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    return fig


def create_organizer_chart(df: pd.DataFrame) -> Optional[plt.Figure]:
    """
    Create a horizontal bar chart of top organizers.

    Args:
        df: DataFrame with organizer data

    Returns:
        matplotlib Figure object
    """
    setup_style()

    if df.empty:
        return None

    # Take top 8 for readability
    df = df.head(8).iloc[::-1]  # Reverse for horizontal bar chart

    fig, ax = plt.subplots(figsize=(10, 6))

    # Truncate long organizer names
    organizers = df['organizer'].apply(lambda x: x[:35] + '...' if len(x) > 35 else x)

    bars = ax.barh(
        organizers,
        df['total_hours'],
        color=COLORS['secondary'],
        edgecolor='white',
        linewidth=1
    )

    # Add value labels
    for bar, meetings in zip(bars, df['meetings']):
        width = bar.get_width()
        ax.annotate(
            f'{width:.1f}h ({meetings} mtgs)',
            xy=(width, bar.get_y() + bar.get_height() / 2),
            xytext=(3, 0),
            textcoords="offset points",
            ha='left',
            va='center',
            fontsize=9
        )

    ax.set_xlabel('Total Hours')
    ax.set_title('Top Meeting Organizers by Time')

    # Clean up axes
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    return fig
