"""
Metrics calculations for calendar audit.
Computes KPIs, analysis tables, and pattern detection.
"""

import pandas as pd
from typing import Dict, Any
from collections import Counter


def calculate_kpis(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Calculate key performance indicators for the dashboard.

    Returns dict with:
    - total_hours: Total meeting hours
    - total_meetings: Count of meetings
    - avg_duration: Average meeting duration in minutes
    - recurring_pct: Estimated percentage of time in recurring meetings
    """
    if df.empty:
        return {
            'total_hours': 0,
            'total_meetings': 0,
            'avg_duration': 0,
            'recurring_pct': 0,
        }

    total_minutes = df['duration_minutes'].sum()
    total_meetings = len(df)
    avg_duration = df['duration_minutes'].mean()

    # Heuristic for recurring meetings:
    # If the same subject appears 2+ times, consider it recurring
    recurring_minutes = estimate_recurring_time(df)
    recurring_pct = (recurring_minutes / total_minutes * 100) if total_minutes > 0 else 0

    return {
        'total_hours': round(total_minutes / 60, 1),
        'total_meetings': total_meetings,
        'avg_duration': round(avg_duration, 1),
        'recurring_pct': round(recurring_pct, 1),
    }


def estimate_recurring_time(df: pd.DataFrame) -> float:
    """
    Estimate time spent in recurring meetings using heuristics.

    Heuristics used:
    1. Meetings with identical subjects appearing 2+ times
    2. Meetings with common recurring patterns in subject (Weekly, Daily, Standup, etc.)
    """
    if df.empty:
        return 0

    # Normalize subjects for comparison
    subjects = df['subject'].str.lower().str.strip()

    # Count subject occurrences
    subject_counts = subjects.value_counts()

    # Find subjects that appear multiple times
    recurring_subjects = subject_counts[subject_counts >= 2].index.tolist()

    # Common recurring meeting keywords
    recurring_keywords = [
        'weekly', 'daily', 'standup', 'stand-up', 'stand up',
        'sync', '1:1', '1-1', 'one on one', 'recurring',
        'monday', 'tuesday', 'wednesday', 'thursday', 'friday',
        'team meeting', 'staff meeting', 'check-in', 'check in',
        'retro', 'sprint', 'scrum', 'planning', 'review'
    ]

    # Create mask for recurring meetings
    is_recurring = subjects.isin(recurring_subjects)

    # Also flag meetings with recurring keywords
    for keyword in recurring_keywords:
        is_recurring = is_recurring | subjects.str.contains(keyword, na=False)

    # Sum duration of recurring meetings
    recurring_minutes = df.loc[is_recurring, 'duration_minutes'].sum()

    return recurring_minutes


def get_top_meetings_by_time(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """
    Get meetings that consume the most total time (aggregated by subject).

    Returns DataFrame with: subject, occurrences, total_hours, avg_duration
    """
    if df.empty:
        return pd.DataFrame(columns=['subject', 'occurrences', 'total_hours', 'avg_duration'])

    grouped = df.groupby('subject').agg(
        occurrences=('subject', 'count'),
        total_minutes=('duration_minutes', 'sum'),
        avg_duration=('duration_minutes', 'mean')
    ).reset_index()

    grouped['total_hours'] = (grouped['total_minutes'] / 60).round(1)
    grouped['avg_duration'] = grouped['avg_duration'].round(0).astype(int)

    grouped = grouped.sort_values('total_hours', ascending=False).head(top_n)

    return grouped[['subject', 'occurrences', 'total_hours', 'avg_duration']]


def get_top_organizers(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """
    Get organizers who schedule the most meeting time.

    Returns DataFrame with: organizer, meetings, total_hours
    """
    if df.empty or df['organizer'].str.strip().eq('').all():
        return pd.DataFrame(columns=['organizer', 'meetings', 'total_hours'])

    # Filter out empty organizers
    has_organizer = df['organizer'].str.strip() != ''
    df_with_org = df[has_organizer]

    if df_with_org.empty:
        return pd.DataFrame(columns=['organizer', 'meetings', 'total_hours'])

    grouped = df_with_org.groupby('organizer').agg(
        meetings=('organizer', 'count'),
        total_minutes=('duration_minutes', 'sum')
    ).reset_index()

    grouped['total_hours'] = (grouped['total_minutes'] / 60).round(1)
    grouped = grouped.sort_values('total_hours', ascending=False).head(top_n)

    return grouped[['organizer', 'meetings', 'total_hours']]


def get_long_meetings(df: pd.DataFrame, threshold_minutes: int = 60) -> pd.DataFrame:
    """
    Get meetings longer than the threshold.

    Returns DataFrame with: subject, duration_minutes, date, organizer
    """
    if df.empty:
        return pd.DataFrame(columns=['subject', 'duration_minutes', 'date', 'organizer'])

    long_meetings = df[df['duration_minutes'] > threshold_minutes].copy()
    long_meetings = long_meetings.sort_values('duration_minutes', ascending=False)

    # Format for display
    result = long_meetings[['subject', 'duration_minutes', 'date', 'organizer']].copy()
    result['duration_minutes'] = result['duration_minutes'].astype(int)

    return result.head(20)


def get_weekday_distribution(df: pd.DataFrame) -> pd.DataFrame:
    """
    Get meeting time distribution by weekday.

    Returns DataFrame with: weekday, total_hours, meeting_count
    """
    if df.empty:
        return pd.DataFrame(columns=['weekday', 'total_hours', 'meeting_count'])

    # Order weekdays properly
    weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

    grouped = df.groupby(['weekday', 'weekday_num']).agg(
        total_minutes=('duration_minutes', 'sum'),
        meeting_count=('subject', 'count')
    ).reset_index()

    grouped['total_hours'] = (grouped['total_minutes'] / 60).round(1)
    grouped = grouped.sort_values('weekday_num')

    return grouped[['weekday', 'total_hours', 'meeting_count']]


def get_duration_distribution(df: pd.DataFrame) -> pd.DataFrame:
    """
    Get distribution of meeting durations for histogram.

    Returns DataFrame with duration data.
    """
    if df.empty:
        return pd.DataFrame(columns=['duration_minutes'])

    return df[['duration_minutes']].copy()


def get_daily_load(df: pd.DataFrame) -> pd.DataFrame:
    """
    Get meeting load per day.

    Returns DataFrame with: date, total_hours, meeting_count
    """
    if df.empty:
        return pd.DataFrame(columns=['date', 'total_hours', 'meeting_count'])

    grouped = df.groupby('date').agg(
        total_minutes=('duration_minutes', 'sum'),
        meeting_count=('subject', 'count')
    ).reset_index()

    grouped['total_hours'] = (grouped['total_minutes'] / 60).round(1)
    grouped = grouped.sort_values('date')

    return grouped[['date', 'total_hours', 'meeting_count']]


def detect_patterns(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Detect patterns in meeting data for insights.

    Returns dict with various pattern metrics.
    """
    if df.empty:
        return {}

    patterns = {}

    # Busiest day of week
    weekday_dist = get_weekday_distribution(df)
    if not weekday_dist.empty:
        busiest = weekday_dist.loc[weekday_dist['total_hours'].idxmax()]
        patterns['busiest_day'] = busiest['weekday']
        patterns['busiest_day_hours'] = busiest['total_hours']

    # Meeting duration insights
    patterns['short_meetings'] = len(df[df['duration_minutes'] <= 30])
    patterns['medium_meetings'] = len(df[(df['duration_minutes'] > 30) & (df['duration_minutes'] <= 60)])
    patterns['long_meetings'] = len(df[df['duration_minutes'] > 60])

    # Time in long meetings
    long_meeting_hours = df[df['duration_minutes'] > 60]['duration_minutes'].sum() / 60
    patterns['long_meeting_hours'] = round(long_meeting_hours, 1)

    # Average meetings per day
    daily_load = get_daily_load(df)
    if not daily_load.empty:
        patterns['avg_meetings_per_day'] = round(daily_load['meeting_count'].mean(), 1)
        patterns['avg_hours_per_day'] = round(daily_load['total_hours'].mean(), 1)
        patterns['max_meetings_day'] = daily_load['meeting_count'].max()

    # Most common meeting duration (rounded to 15 min)
    rounded_durations = (df['duration_minutes'] / 15).round() * 15
    most_common = rounded_durations.mode()
    if len(most_common) > 0:
        patterns['most_common_duration'] = int(most_common.iloc[0])

    # Early morning / late evening meetings
    if 'start_datetime' in df.columns:
        hours = df['start_datetime'].apply(lambda x: x.hour if x else 12)
        patterns['early_meetings'] = len(df[hours < 9])  # Before 9 AM
        patterns['late_meetings'] = len(df[hours >= 17])  # 5 PM or later

    return patterns
