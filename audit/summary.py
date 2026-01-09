"""
Executive summary generator for calendar audit.
Uses EA (Executive Assistant) voice - professional, supportive, non-directive.
"""

import pandas as pd
from typing import Dict, Any, List
from datetime import datetime


def generate_observations(
    kpis: Dict[str, Any],
    patterns: Dict[str, Any],
    top_meetings: pd.DataFrame,
    top_organizers: pd.DataFrame
) -> List[str]:
    """
    Generate data-driven observations about meeting patterns.
    Uses neutral, observational language.
    """
    observations = []

    # Meeting load observation
    if 'avg_hours_per_day' in patterns:
        avg_daily = patterns['avg_hours_per_day']
        observations.append(
            f"The calendar reflects an average of {avg_daily} hours in meetings per day "
            f"across the analyzed period."
        )

    # Recurring meetings observation
    if kpis['recurring_pct'] > 30:
        observations.append(
            f"Approximately {kpis['recurring_pct']}% of meeting time appears to be "
            "allocated to recurring commitments."
        )

    # Busiest day observation
    if 'busiest_day' in patterns:
        observations.append(
            f"{patterns['busiest_day']} currently carries the highest meeting load "
            f"at {patterns['busiest_day_hours']} hours."
        )

    # Long meetings observation
    if patterns.get('long_meetings', 0) >= 3:
        observations.append(
            f"{patterns['long_meetings']} meetings exceed 60 minutes, "
            f"representing {patterns.get('long_meeting_hours', 0)} hours total."
        )

    # Top time consumer observation
    if not top_meetings.empty:
        top = top_meetings.iloc[0]
        if top['occurrences'] >= 2:
            observations.append(
                f"'{top['subject'][:35]}' accounts for {top['total_hours']} hours "
                f"across {top['occurrences']} occurrences."
            )

    # Meeting duration pattern
    if 'most_common_duration' in patterns:
        observations.append(
            f"The most frequently scheduled meeting length is {patterns['most_common_duration']} minutes."
        )

    return observations[:5]  # Cap at 5 observations


def generate_considerations(
    kpis: Dict[str, Any],
    patterns: Dict[str, Any],
    top_meetings: pd.DataFrame,
    long_meetings: pd.DataFrame
) -> List[str]:
    """
    Generate suggested considerations using EA voice.
    Frames as options, not directives.
    """
    considerations = []

    # Meeting duration consideration
    if 'most_common_duration' in patterns:
        if patterns['most_common_duration'] == 60:
            considerations.append(
                "You may want to consider whether 50-minute defaults would create "
                "helpful buffer time between sessions."
            )
        elif patterns['most_common_duration'] == 30:
            considerations.append(
                "One potential adjustment is shifting to 25-minute meetings where appropriate, "
                "allowing brief transitions between calls."
            )

    # Recurring meeting consideration
    if kpis['recurring_pct'] > 40:
        considerations.append(
            "Based on the recurring meeting volume, a periodic review of standing "
            "commitments may surface opportunities to consolidate or adjust frequency."
        )

    # Long meetings consideration
    if patterns.get('long_meetings', 0) >= 3:
        considerations.append(
            "For longer sessions, there may be value in exploring whether pre-reads "
            "or async updates could reduce required meeting time."
        )

    # Top meeting consideration
    if not top_meetings.empty:
        top = top_meetings.iloc[0]
        if top['occurrences'] >= 4:
            considerations.append(
                f"The frequency of '{top['subject'][:30]}' could be worth revisiting "
                "to confirm the current cadence still aligns with priorities."
            )

    # Busiest day consideration
    if 'busiest_day' in patterns and patterns.get('busiest_day_hours', 0) > 5:
        considerations.append(
            f"Given the concentration of meetings on {patterns['busiest_day']}, "
            "protecting a focus block that day may be beneficial."
        )

    # High meeting volume consideration
    if patterns.get('avg_meetings_per_day', 0) > 5:
        considerations.append(
            "With the current meeting density, there may be opportunities to identify "
            "sessions where a delegate or summary could serve as an alternative."
        )

    return considerations[:5]  # Cap at 5 considerations


def generate_overview(kpis: Dict[str, Any], patterns: Dict[str, Any]) -> str:
    """
    Generate a brief 2-3 sentence overview of calendar patterns.
    """
    total_hours = kpis['total_hours']
    total_meetings = kpis['total_meetings']
    avg_duration = int(kpis['avg_duration'])

    overview = f"This analysis covers {total_meetings} meetings totaling {total_hours} hours, "
    overview += f"with an average duration of {avg_duration} minutes. "

    if patterns.get('avg_hours_per_day', 0) > 5:
        overview += "The current pattern reflects a meeting-intensive schedule. "
    elif patterns.get('avg_hours_per_day', 0) > 3:
        overview += "The calendar shows a moderate level of meeting activity. "
    else:
        overview += "Meeting load appears balanced relative to available time. "

    if kpis['recurring_pct'] > 40:
        overview += f"Recurring meetings represent a notable portion ({kpis['recurring_pct']}%) of total time."

    return overview


def generate_closing(patterns: Dict[str, Any]) -> str:
    """
    Generate a brief closing paragraph focused on supporting priorities.
    """
    closing = "These observations are intended to support informed decisions about calendar management. "

    if patterns.get('avg_hours_per_day', 0) > 4:
        closing += "Small adjustments to meeting frequency or duration can often create meaningful "
        closing += "capacity for focused work. "
    else:
        closing += "The current structure provides a foundation that can be refined as priorities evolve. "

    closing += "Happy to discuss any of these patterns in more detail or explore specific adjustments."

    return closing


def generate_full_summary(
    kpis: Dict[str, Any],
    patterns: Dict[str, Any],
    top_meetings: pd.DataFrame,
    top_organizers: pd.DataFrame,
    long_meetings: pd.DataFrame
) -> str:
    """
    Generate the complete executive summary in EA voice.
    Designed to fit on one page.
    """
    overview = generate_overview(kpis, patterns)
    observations = generate_observations(kpis, patterns, top_meetings, top_organizers)
    considerations = generate_considerations(kpis, patterns, top_meetings, long_meetings)
    closing = generate_closing(patterns)

    # Build markdown document (concise, one-page format)
    lines = [
        "# Calendar Audit Summary",
        "",
        f"*Prepared {datetime.now().strftime('%B %d, %Y')}*",
        "",
        "---",
        "",
        "## Overview",
        "",
        overview,
        "",
        "---",
        "",
        "## Key Observations",
        "",
    ]

    for obs in observations:
        lines.append(f"- {obs}")

    lines.extend([
        "",
        "---",
        "",
        "## Considerations",
        "",
    ])

    for consideration in considerations:
        lines.append(f"- {consideration}")

    lines.extend([
        "",
        "---",
        "",
        closing,
        "",
    ])

    return "\n".join(lines)
