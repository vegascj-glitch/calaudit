"""Quick test script to process sample data and show results."""

from audit.parser import parse_calendar_file, apply_filters
from audit.metrics import (
    calculate_kpis,
    get_top_meetings_by_time,
    get_top_organizers,
    get_long_meetings,
    get_weekday_distribution,
    detect_patterns,
)
from audit.summary import generate_full_summary

def test_sample(filename: str):
    print(f"\n{'='*60}")
    print(f"Testing: {filename}")
    print('='*60)

    # Read file
    with open(filename, 'rb') as f:
        content = f.read()

    # Parse
    df, source, warnings = parse_calendar_file(content, filename)
    print(f"\nDetected source: {source}")
    print(f"Total events loaded: {len(df)}")

    if warnings:
        print(f"Warnings: {warnings}")

    # Apply filters (exclude all-day events)
    filtered = apply_filters(df, exclude_all_day=True)
    print(f"Events after filtering: {len(filtered)}")

    # Calculate KPIs
    kpis = calculate_kpis(filtered)
    print(f"\n--- KPIs ---")
    print(f"Total Meeting Hours: {kpis['total_hours']}h")
    print(f"Total Meetings: {kpis['total_meetings']}")
    print(f"Average Duration: {int(kpis['avg_duration'])} minutes")
    print(f"Recurring Meeting %: {kpis['recurring_pct']}%")

    # Patterns
    patterns = detect_patterns(filtered)
    print(f"\n--- Patterns ---")
    print(f"Busiest Day: {patterns.get('busiest_day', 'N/A')} ({patterns.get('busiest_day_hours', 0)}h)")
    print(f"Avg Meetings/Day: {patterns.get('avg_meetings_per_day', 'N/A')}")
    print(f"Most Common Duration: {patterns.get('most_common_duration', 'N/A')} min")

    # Top meetings
    top_meetings = get_top_meetings_by_time(filtered, top_n=5)
    print(f"\n--- Top 5 Meetings by Time ---")
    for _, row in top_meetings.iterrows():
        print(f"  {row['subject'][:40]:<40} | {row['total_hours']}h ({row['occurrences']}x)")

    # Top organizers
    top_org = get_top_organizers(filtered, top_n=5)
    if not top_org.empty:
        print(f"\n--- Top 5 Organizers ---")
        for _, row in top_org.iterrows():
            print(f"  {row['organizer'][:35]:<35} | {row['total_hours']}h ({row['meetings']} mtgs)")

    # Weekday distribution
    weekday = get_weekday_distribution(filtered)
    print(f"\n--- Meeting Hours by Weekday ---")
    for _, row in weekday.iterrows():
        bar = '#' * int(row['total_hours'])
        print(f"  {row['weekday']:<10} | {bar} {row['total_hours']}h ({row['meeting_count']} mtgs)")

    # Long meetings
    long = get_long_meetings(filtered, threshold_minutes=60)
    print(f"\n--- Long Meetings (>60 min): {len(long)} found ---")
    for _, row in long.head(5).iterrows():
        print(f"  {row['subject'][:40]:<40} | {int(row['duration_minutes'])} min")

    # Generate summary
    summary = generate_full_summary(
        kpis=kpis,
        patterns=patterns,
        top_meetings=top_meetings,
        top_organizers=top_org,
        long_meetings=long
    )

    print(f"\n{'='*60}")
    print("EXECUTIVE SUMMARY")
    print('='*60)
    print(summary)

    return summary


if __name__ == "__main__":
    # Test Outlook CSV sample
    test_sample("sample_data/sample_outlook.csv")

    print("\n" + "="*80 + "\n")

    # Test Google CSV sample
    test_sample("sample_data/sample_google.csv")

    print("\n" + "="*80 + "\n")

    # Test Google ICS sample
    test_sample("sample_data/sample_google.ics")
