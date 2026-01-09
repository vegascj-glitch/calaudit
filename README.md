# Calendar Audit Demo

A Streamlit application for Executive Assistants to analyze calendar exports from Outlook and Google Calendar. Identify meeting patterns, time leaks, and generate executive-ready recommendations.

## Features

- **Single file upload** - Upload ONE .csv or .ics file at a time
- **Auto-detect format** - Supports Outlook CSV and Google Calendar ICS
- **KPI Dashboard** - Total hours, meeting count, average duration, recurring percentage
- **Visual Analytics** - Meeting load by weekday, duration distribution charts
- **Analysis Tables** - Top meetings by time, top organizers, long meetings list
- **Executive Summary** - Rule-based insights and actionable recommendations
- **Export** - Download summary as Markdown for emails or presentations

## Quick Start

### 1. Install Dependencies

```bash
cd calendar_audit_demo
pip install -r requirements.txt
```

### 2. Run the Application

```bash
streamlit run app.py
```

### 3. Upload a Calendar File

Upload ONE file at a time: either a .csv (Outlook) or .ics (Google Calendar).

## How to Export Your Calendar

### Outlook (CSV)

1. Open Outlook Calendar
2. Go to **File > Open & Export > Import/Export**
3. Select **Export to a file**
4. Choose **Comma Separated Values**
5. Select your calendar folder
6. Save the .csv file
7. **Upload the .csv file** to the app

### Google Calendar (ICS)

1. Go to [Google Calendar Settings](https://calendar.google.com/calendar/r/settings/export)
2. Click **Export** (downloads a .zip file)
3. **Unzip** the downloaded file
4. **Upload ONE .ics file** to the app (each .ics = one calendar)

> **Note:** Google exports all calendars in a single .zip. After unzipping, choose the specific calendar (.ics file) you want to analyze.

## Sample Data

Sample files are included for testing:

- `sample_data/sample_outlook.csv` - Outlook CSV format
- `sample_data/sample_google.csv` - Google Calendar CSV format

## Project Structure

```
calendar_audit_demo/
├── app.py                 # Main Streamlit application
├── audit/
│   ├── __init__.py
│   ├── parser.py          # CSV and ICS parsing
│   ├── metrics.py         # KPI and analysis calculations
│   ├── charts.py          # Matplotlib visualizations
│   └── summary.py         # Executive summary generation
├── sample_data/
│   ├── sample_outlook.csv
│   └── sample_google.csv
├── requirements.txt
└── README.md
```

## Filters

The sidebar provides controls to customize your analysis:

- **Exclude all-day events** - Remove holidays, PTO, blocked days (default: ON)
- **Minimum duration** - Exclude short meetings below a threshold
- **Keyword exclusion** - Remove meetings by subject (e.g., "lunch, blocked, hold")

## Metrics Explained

| Metric | Description |
|--------|-------------|
| Total Meeting Hours | Sum of all meeting durations |
| Total Meetings | Count of meetings after filters |
| Average Duration | Mean meeting length in minutes |
| Recurring % | Estimated time in recurring meetings (heuristic-based) |

## Executive Summary

The generated summary includes:

1. **Key Insights** - 4-6 data-driven observations about meeting patterns
2. **Recommended Actions** - 4-6 actionable suggestions (shorten, decline, consolidate, delegate)
3. **Executive Paragraph** - Professional summary suitable for leadership communication

## Tech Stack

- **Python 3.8+**
- **Streamlit** - Web UI framework
- **pandas** - Data processing
- **matplotlib** - Visualizations
- **icalendar** - ICS file parsing

## V1 Limitations

- Single file upload only (no batch/zip upload)
- One calendar per analysis session
- No authentication required

## License

MIT License - Free for demo and internal use.
