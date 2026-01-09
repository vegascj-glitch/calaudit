"""
Calendar Parser for CSV and ICS Exports
Handles Outlook CSV, Google Calendar CSV, and ICS (iCalendar) formats.
"""

import pandas as pd
from datetime import datetime, date
from typing import Tuple, Optional


# Column mappings for different calendar sources
OUTLOOK_COLUMNS = {
    'subject': ['Subject'],
    'start_date': ['Start Date'],
    'start_time': ['Start Time'],
    'end_date': ['End Date'],
    'end_time': ['End Time'],
    'all_day': ['All day event', 'All Day Event'],
    'organizer': ['Organizer', 'Meeting Organizer'],
    'attendees': ['Required Attendees', 'Attendees'],
    'optional_attendees': ['Optional Attendees'],
    'location': ['Location'],
}

GOOGLE_COLUMNS = {
    'subject': ['Subject', 'Title'],
    'start_date': ['Start Date'],
    'start_time': ['Start Time'],
    'end_date': ['End Date'],
    'end_time': ['End Time'],
    'all_day': ['All Day Event', 'All day event'],
    'description': ['Description'],
    'location': ['Location'],
}

# Signature columns that help identify the source
OUTLOOK_SIGNATURE = ['Organizer', 'Required Attendees', 'Meeting Organizer']
GOOGLE_SIGNATURE = ['Description', 'Private']


def detect_file_type(file_content: bytes, filename: str = '') -> str:
    """Detect if file is CSV or ICS based on content and filename."""
    filename_lower = filename.lower()
    if filename_lower.endswith('.ics'):
        return 'ics'
    if filename_lower.endswith('.csv'):
        return 'csv'

    # Try to detect from content
    try:
        content_start = file_content[:500].decode('utf-8', errors='ignore').strip()
        if content_start.startswith('BEGIN:VCALENDAR'):
            return 'ics'
    except:
        pass

    return 'csv'  # Default to CSV


def detect_csv_source(df: pd.DataFrame) -> str:
    """
    Auto-detect whether CSV is from Outlook or Google Calendar.
    Returns: 'outlook', 'google', or 'unknown'
    """
    columns = set(df.columns.str.strip())

    outlook_matches = sum(1 for col in OUTLOOK_SIGNATURE if col in columns)
    google_matches = sum(1 for col in GOOGLE_SIGNATURE if col in columns)

    if outlook_matches > google_matches:
        return 'outlook'
    elif google_matches > outlook_matches:
        return 'google'
    elif 'Subject' in columns and 'Start Date' in columns:
        return 'google'
    else:
        return 'unknown'


def find_column(df: pd.DataFrame, possible_names: list) -> Optional[str]:
    """Find the first matching column name from a list of possibilities."""
    columns = df.columns.str.strip().tolist()
    for name in possible_names:
        if name in columns:
            return name
    return None


def parse_datetime_str(date_str: str, time_str: str = None) -> Optional[datetime]:
    """Parse date and optional time strings into a datetime object."""
    if pd.isna(date_str) or str(date_str).strip() == '':
        return None

    date_str = str(date_str).strip()
    time_str = str(time_str).strip() if time_str and not pd.isna(time_str) else ''

    date_formats = [
        '%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%m-%d-%Y',
        '%m/%d/%y', '%d/%m/%y', '%Y/%m/%d'
    ]
    time_formats = ['%I:%M:%S %p', '%I:%M %p', '%H:%M:%S', '%H:%M', '']

    for date_fmt in date_formats:
        for time_fmt in time_formats:
            try:
                if time_str and time_fmt:
                    full_fmt = f"{date_fmt} {time_fmt}"
                    full_str = f"{date_str} {time_str}"
                else:
                    full_fmt = date_fmt
                    full_str = date_str
                return datetime.strptime(full_str.strip(), full_fmt)
            except ValueError:
                continue
    return None


def parse_bool(value) -> bool:
    """Parse various boolean representations."""
    if pd.isna(value):
        return False
    value = str(value).lower().strip()
    return value in ('true', 'yes', '1', 'on')


def normalize_outlook(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize Outlook CSV to common schema."""
    normalized = pd.DataFrame()

    subject_col = find_column(df, OUTLOOK_COLUMNS['subject'])
    start_date_col = find_column(df, OUTLOOK_COLUMNS['start_date'])
    start_time_col = find_column(df, OUTLOOK_COLUMNS['start_time'])
    end_date_col = find_column(df, OUTLOOK_COLUMNS['end_date'])
    end_time_col = find_column(df, OUTLOOK_COLUMNS['end_time'])
    all_day_col = find_column(df, OUTLOOK_COLUMNS['all_day'])
    organizer_col = find_column(df, OUTLOOK_COLUMNS['organizer'])
    attendees_col = find_column(df, OUTLOOK_COLUMNS['attendees'])
    optional_col = find_column(df, OUTLOOK_COLUMNS['optional_attendees'])
    location_col = find_column(df, OUTLOOK_COLUMNS['location'])

    normalized['subject'] = df[subject_col].fillna('(No Subject)') if subject_col else '(No Subject)'

    start_times = []
    end_times = []
    for idx, row in df.iterrows():
        start_dt = parse_datetime_str(
            row.get(start_date_col, '') if start_date_col else '',
            row.get(start_time_col, '') if start_time_col else ''
        )
        end_dt = parse_datetime_str(
            row.get(end_date_col, '') if end_date_col else '',
            row.get(end_time_col, '') if end_time_col else ''
        )
        start_times.append(start_dt)
        end_times.append(end_dt)

    normalized['start_datetime'] = start_times
    normalized['end_datetime'] = end_times

    if all_day_col:
        normalized['is_all_day'] = df[all_day_col].apply(parse_bool)
    else:
        normalized['is_all_day'] = False

    normalized['organizer'] = df[organizer_col].fillna('') if organizer_col else ''

    if attendees_col and optional_col:
        normalized['attendees'] = (
            df[attendees_col].fillna('') + '; ' + df[optional_col].fillna('')
        ).str.strip('; ')
    elif attendees_col:
        normalized['attendees'] = df[attendees_col].fillna('')
    else:
        normalized['attendees'] = ''

    normalized['location'] = df[location_col].fillna('') if location_col else ''

    return normalized


def normalize_google_csv(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize Google Calendar CSV to common schema."""
    normalized = pd.DataFrame()

    subject_col = find_column(df, GOOGLE_COLUMNS['subject'])
    start_date_col = find_column(df, GOOGLE_COLUMNS['start_date'])
    start_time_col = find_column(df, GOOGLE_COLUMNS['start_time'])
    end_date_col = find_column(df, GOOGLE_COLUMNS['end_date'])
    end_time_col = find_column(df, GOOGLE_COLUMNS['end_time'])
    all_day_col = find_column(df, GOOGLE_COLUMNS['all_day'])
    location_col = find_column(df, GOOGLE_COLUMNS['location'])

    normalized['subject'] = df[subject_col].fillna('(No Subject)') if subject_col else '(No Subject)'

    start_times = []
    end_times = []
    for idx, row in df.iterrows():
        start_dt = parse_datetime_str(
            row.get(start_date_col, '') if start_date_col else '',
            row.get(start_time_col, '') if start_time_col else ''
        )
        end_dt = parse_datetime_str(
            row.get(end_date_col, '') if end_date_col else '',
            row.get(end_time_col, '') if end_time_col else ''
        )
        start_times.append(start_dt)
        end_times.append(end_dt)

    normalized['start_datetime'] = start_times
    normalized['end_datetime'] = end_times

    if all_day_col:
        normalized['is_all_day'] = df[all_day_col].apply(parse_bool)
    else:
        normalized['is_all_day'] = False

    normalized['organizer'] = ''
    normalized['attendees'] = ''
    normalized['location'] = df[location_col].fillna('') if location_col else ''

    return normalized


def parse_ics_file(file_content: bytes) -> Tuple[pd.DataFrame, list]:
    """
    Parse an ICS (iCalendar) file into normalized DataFrame.

    Returns:
        - Normalized DataFrame
        - List of warnings
    """
    warnings = []

    try:
        from icalendar import Calendar
    except ImportError:
        warnings.append("icalendar library not installed. Run: pip install icalendar")
        return pd.DataFrame(), warnings

    # Decode content
    content = None
    for encoding in ['utf-8', 'latin-1', 'cp1252']:
        try:
            content = file_content.decode(encoding)
            break
        except UnicodeDecodeError:
            continue

    if content is None:
        warnings.append("Could not decode ICS file")
        return pd.DataFrame(), warnings

    # Parse calendar
    try:
        cal = Calendar.from_ical(content)
    except Exception as e:
        warnings.append(f"Error parsing ICS file: {str(e)}")
        return pd.DataFrame(), warnings

    # Extract events
    events = []
    for component in cal.walk():
        if component.name == 'VEVENT':
            event = extract_ics_event(component)
            if event:
                events.append(event)

    if not events:
        warnings.append("No events found in ICS file")
        return pd.DataFrame(), warnings

    df = pd.DataFrame(events)
    return df, warnings


def extract_ics_event(component) -> Optional[dict]:
    """Extract event data from an ICS VEVENT component."""
    try:
        # Get summary/subject
        subject = str(component.get('summary', '(No Subject)') or '(No Subject)')

        # Get start time
        dtstart = component.get('dtstart')
        if dtstart is None:
            return None
        start_val = dtstart.dt

        # Get end time
        dtend = component.get('dtend')
        if dtend:
            end_val = dtend.dt
        else:
            # If no end, use duration or same as start
            duration = component.get('duration')
            if duration:
                end_val = start_val + duration.dt
            else:
                end_val = start_val

        # Check if all-day event (date vs datetime)
        is_all_day = isinstance(start_val, date) and not isinstance(start_val, datetime)

        # Convert date to datetime if needed
        if is_all_day:
            start_datetime = datetime.combine(start_val, datetime.min.time())
            if isinstance(end_val, date) and not isinstance(end_val, datetime):
                end_datetime = datetime.combine(end_val, datetime.min.time())
            else:
                end_datetime = end_val
        else:
            # Handle timezone-aware datetimes
            start_datetime = start_val.replace(tzinfo=None) if hasattr(start_val, 'replace') else start_val
            end_datetime = end_val.replace(tzinfo=None) if hasattr(end_val, 'replace') else end_val

        # Ensure both are datetime objects
        if isinstance(start_datetime, date) and not isinstance(start_datetime, datetime):
            start_datetime = datetime.combine(start_datetime, datetime.min.time())
        if isinstance(end_datetime, date) and not isinstance(end_datetime, datetime):
            end_datetime = datetime.combine(end_datetime, datetime.min.time())

        # Get organizer
        organizer = component.get('organizer')
        organizer_str = ''
        if organizer:
            organizer_str = str(organizer).replace('mailto:', '')

        # Get attendees
        attendees = component.get('attendee')
        attendees_str = ''
        if attendees:
            if isinstance(attendees, list):
                attendees_str = '; '.join(str(a).replace('mailto:', '') for a in attendees)
            else:
                attendees_str = str(attendees).replace('mailto:', '')

        # Get location
        location = str(component.get('location', '') or '')

        return {
            'subject': subject,
            'start_datetime': start_datetime,
            'end_datetime': end_datetime,
            'is_all_day': is_all_day,
            'organizer': organizer_str,
            'attendees': attendees_str,
            'location': location,
        }

    except Exception as e:
        return None


def calculate_derived_fields(df: pd.DataFrame) -> pd.DataFrame:
    """Add derived fields to normalized data."""
    df['duration_minutes'] = df.apply(
        lambda row: (row['end_datetime'] - row['start_datetime']).total_seconds() / 60
        if row['start_datetime'] and row['end_datetime'] else 0,
        axis=1
    )

    df['weekday'] = df['start_datetime'].apply(
        lambda dt: dt.strftime('%A') if dt else 'Unknown'
    )

    df['weekday_num'] = df['start_datetime'].apply(
        lambda dt: dt.weekday() if dt else 7
    )

    df['date'] = df['start_datetime'].apply(
        lambda dt: dt.date() if dt else None
    )

    return df


def parse_calendar_file(
    file_content: bytes,
    filename: str = '',
    source_override: Optional[str] = None
) -> Tuple[pd.DataFrame, str, list]:
    """
    Main parsing function. Detects file type, parses, and normalizes data.

    Args:
        file_content: Raw file bytes
        filename: Original filename (used for type detection)
        source_override: Optional override for CSV source type

    Returns:
        - Normalized DataFrame
        - Detected source ('outlook', 'google', 'ics', 'unknown')
        - List of any warnings/errors encountered
    """
    warnings = []

    # Detect file type
    file_type = detect_file_type(file_content, filename)

    if file_type == 'ics':
        # Parse ICS file
        df, ics_warnings = parse_ics_file(file_content)
        warnings.extend(ics_warnings)

        if df.empty:
            return pd.DataFrame(), 'ics', warnings

        source = 'ics'

    else:
        # Parse CSV file
        for encoding in ['utf-8', 'latin-1', 'cp1252']:
            try:
                from io import StringIO
                content = file_content.decode(encoding)
                df = pd.read_csv(StringIO(content))
                break
            except UnicodeDecodeError:
                continue
            except Exception as e:
                warnings.append(f"Error reading CSV: {str(e)}")
                return pd.DataFrame(), 'unknown', warnings
        else:
            warnings.append("Could not decode CSV file with any supported encoding")
            return pd.DataFrame(), 'unknown', warnings

        if df.empty:
            warnings.append("CSV file is empty")
            return pd.DataFrame(), 'unknown', warnings

        df.columns = df.columns.str.strip()

        detected_source = detect_csv_source(df)
        source = source_override if source_override else detected_source

        if source != detected_source and source_override:
            warnings.append(f"Using manual override: {source} (auto-detected: {detected_source})")

        if source == 'outlook':
            df = normalize_outlook(df)
        elif source == 'google':
            df = normalize_google_csv(df)
        else:
            warnings.append("Could not detect calendar source. Attempting Google format.")
            df = normalize_google_csv(df)
            source = 'google'

    # Remove rows with invalid datetimes
    initial_count = len(df)
    df = df.dropna(subset=['start_datetime', 'end_datetime'])
    dropped = initial_count - len(df)
    if dropped > 0:
        warnings.append(f"Dropped {dropped} rows with invalid dates")

    # Calculate derived fields
    df = calculate_derived_fields(df)

    # Filter out negative durations
    invalid_duration = df['duration_minutes'] < 0
    if invalid_duration.any():
        count = invalid_duration.sum()
        warnings.append(f"Dropped {count} rows with invalid duration")
        df = df[~invalid_duration]

    return df, source, warnings


# Backwards compatibility alias
def parse_calendar_csv(
    file_content: bytes,
    source_override: Optional[str] = None
) -> Tuple[pd.DataFrame, str, list]:
    """Legacy function name for CSV parsing."""
    return parse_calendar_file(file_content, '', source_override)


def apply_filters(
    df: pd.DataFrame,
    exclude_all_day: bool = True,
    min_duration: int = 0,
    exclude_keywords: list = None
) -> pd.DataFrame:
    """Apply user-specified filters to the normalized data."""
    filtered = df.copy()

    if exclude_all_day:
        filtered = filtered[~filtered['is_all_day']]

    if min_duration > 0:
        filtered = filtered[filtered['duration_minutes'] >= min_duration]

    if exclude_keywords:
        keywords = [k.strip().lower() for k in exclude_keywords if k.strip()]
        if keywords:
            pattern = '|'.join(keywords)
            mask = ~filtered['subject'].str.lower().str.contains(pattern, na=False)
            filtered = filtered[mask]

    return filtered.reset_index(drop=True)
