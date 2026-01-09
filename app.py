"""
Calendar Audit Demo Application

A Streamlit application for analyzing calendar exports from Outlook and Google Calendar.
Provides KPIs, visualizations, and executive-ready summaries for identifying time leaks
and optimization opportunities.

Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd

from audit.parser import parse_calendar_file, apply_filters
from audit.metrics import (
    calculate_kpis,
    get_top_meetings_by_time,
    get_top_organizers,
    get_long_meetings,
    get_weekday_distribution,
    get_duration_distribution,
    get_daily_load,
    detect_patterns,
)
from audit.charts import (
    create_weekday_chart,
    create_duration_histogram,
    create_top_meetings_chart,
    create_organizer_chart,
)
from audit.summary import generate_full_summary


# Page configuration
st.set_page_config(
    page_title="Calendar Audit",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #4A90A4;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #666;
    }
    .section-header {
        border-bottom: 2px solid #4A90A4;
        padding-bottom: 0.5rem;
        margin-bottom: 1rem;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffc107;
        border-radius: 5px;
        padding: 1rem;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #28a745;
        border-radius: 5px;
        padding: 1rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)


def display_kpi_cards(kpis: dict):
    """Display KPI metrics as cards."""
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            label="Total Meeting Hours",
            value=f"{kpis['total_hours']}h",
            help="Total hours spent in meetings"
        )

    with col2:
        st.metric(
            label="Total Meetings",
            value=kpis['total_meetings'],
            help="Number of meetings analyzed"
        )

    with col3:
        st.metric(
            label="Avg Duration",
            value=f"{int(kpis['avg_duration'])}m",
            help="Average meeting duration in minutes"
        )

    with col4:
        st.metric(
            label="Recurring %",
            value=f"{kpis['recurring_pct']}%",
            help="Estimated percentage of time in recurring meetings"
        )


def main():
    # Header
    st.markdown('<p class="main-header">📅 Calendar Audit</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sub-header">Analyze calendar exports to identify meeting patterns, '
        'time leaks, and optimization opportunities</p>',
        unsafe_allow_html=True
    )

    # Sidebar - File Upload and Filters
    with st.sidebar:
        st.header("📁 Data Input")

        uploaded_file = st.file_uploader(
            "Upload Calendar File",
            type=['csv', 'ics'],
            help="Upload ONE file: CSV (Outlook) or ICS (Google Calendar)"
        )

        # Source override (only for CSV files)
        source_override = st.selectbox(
            "CSV Source Override",
            options=['Auto-detect', 'Outlook', 'Google Calendar'],
            index=0,
            help="Override automatic source detection for CSV files (ignored for ICS)"
        )

        st.divider()
        st.header("🎛️ Filters")

        exclude_all_day = st.checkbox(
            "Exclude all-day events",
            value=True,
            help="Remove all-day events from analysis"
        )

        min_duration = st.slider(
            "Minimum meeting duration (minutes)",
            min_value=0,
            max_value=60,
            value=0,
            step=5,
            help="Exclude meetings shorter than this duration"
        )

        exclude_keywords = st.text_input(
            "Exclude by keywords",
            placeholder="lunch, blocked, hold",
            help="Comma-separated keywords to exclude from analysis"
        )

        st.divider()
        st.header("ℹ️ About")
        st.markdown("""
        **Calendar Audit** helps Executive Assistants analyze
        how executives spend time in meetings.

        **Supported formats:**
        - Outlook CSV exports
        - Google Calendar ICS exports

        **Upload one file at a time.**

        **Features:**
        - Auto-detects calendar source
        - KPI dashboard
        - Visual analytics
        - Executive summary generation
        """)

    # Main content
    if uploaded_file is None:
        # Show instructions when no file uploaded
        st.info("👈 Upload a calendar file (.csv or .ics) to get started")

        with st.expander("📖 How to export your calendar", expanded=True):
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Outlook (CSV)")
                st.markdown("""
                1. Open Outlook Calendar
                2. Go to **File → Open & Export → Import/Export**
                3. Select **Export to a file**
                4. Choose **Comma Separated Values**
                5. Select your calendar folder
                6. Save and upload the **.csv** file
                """)

            with col2:
                st.subheader("Google Calendar (ICS)")
                st.markdown("""
                1. Go to [Google Calendar Settings](https://calendar.google.com/calendar/r/settings/export)
                2. Click **Export** (downloads a .zip file)
                3. **Unzip** the downloaded file
                4. Upload **one .ics file** (one calendar)

                *Each .ics file = one calendar*
                """)

        return

    # Process uploaded file
    file_content = uploaded_file.read()
    filename = uploaded_file.name

    # Map source override (only applies to CSV)
    source_map = {
        'Auto-detect': None,
        'Outlook': 'outlook',
        'Google Calendar': 'google'
    }
    override = source_map.get(source_override)

    # Parse file (CSV or ICS)
    with st.spinner("Parsing calendar data..."):
        df, detected_source, warnings = parse_calendar_file(file_content, filename, override)

    # Show warnings
    if warnings:
        for warning in warnings:
            st.warning(warning)

    if df.empty:
        st.error("No valid calendar data found in the uploaded file. Please check the format.")
        return

    # Show detection result
    source_labels = {
        'outlook': 'Outlook (CSV)',
        'google': 'Google Calendar (CSV)',
        'ics': 'ICS (iCalendar)',
        'unknown': 'Unknown'
    }
    source_display = source_labels.get(detected_source, detected_source.title())
    st.success(f"✅ Detected format: **{source_display}** | Found **{len(df)}** events")

    # Apply filters
    keyword_list = [k.strip() for k in exclude_keywords.split(',') if k.strip()] if exclude_keywords else []
    filtered_df = apply_filters(
        df,
        exclude_all_day=exclude_all_day,
        min_duration=min_duration,
        exclude_keywords=keyword_list
    )

    if filtered_df.empty:
        st.warning("No meetings remain after applying filters. Try adjusting filter settings.")
        return

    st.info(f"📊 Analyzing **{len(filtered_df)}** meetings after filters")

    # Calculate metrics
    kpis = calculate_kpis(filtered_df)
    patterns = detect_patterns(filtered_df)
    top_meetings = get_top_meetings_by_time(filtered_df)
    top_organizers = get_top_organizers(filtered_df)
    long_meetings = get_long_meetings(filtered_df)
    weekday_dist = get_weekday_distribution(filtered_df)
    duration_dist = get_duration_distribution(filtered_df)

    # KPI Section
    st.markdown('<h2 class="section-header">📈 Key Metrics</h2>', unsafe_allow_html=True)
    display_kpi_cards(kpis)

    st.divider()

    # Charts Section
    st.markdown('<h2 class="section-header">📊 Visual Analysis</h2>', unsafe_allow_html=True)

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.subheader("Meeting Load by Day")
        weekday_chart = create_weekday_chart(weekday_dist)
        if weekday_chart:
            st.pyplot(weekday_chart)
        else:
            st.info("No data available for weekday chart")

    with chart_col2:
        st.subheader("Duration Distribution")
        duration_chart = create_duration_histogram(filtered_df)
        if duration_chart:
            st.pyplot(duration_chart)
        else:
            st.info("No data available for duration histogram")

    st.divider()

    # Analysis Tables Section
    st.markdown('<h2 class="section-header">📋 Analysis Tables</h2>', unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["Top Meetings", "Top Organizers", "Long Meetings"])

    with tab1:
        st.subheader("Meetings by Total Time")
        if not top_meetings.empty:
            st.dataframe(
                top_meetings,
                column_config={
                    "subject": st.column_config.TextColumn("Meeting", width="large"),
                    "occurrences": st.column_config.NumberColumn("Count", width="small"),
                    "total_hours": st.column_config.NumberColumn("Total Hours", format="%.1f"),
                    "avg_duration": st.column_config.NumberColumn("Avg Duration (min)")
                },
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info("No meeting data available")

    with tab2:
        st.subheader("Organizers by Total Time")
        if not top_organizers.empty:
            st.dataframe(
                top_organizers,
                column_config={
                    "organizer": st.column_config.TextColumn("Organizer", width="large"),
                    "meetings": st.column_config.NumberColumn("Meetings", width="small"),
                    "total_hours": st.column_config.NumberColumn("Total Hours", format="%.1f")
                },
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info("Organizer data not available (may not be included in export)")

    with tab3:
        st.subheader("Meetings Over 60 Minutes")
        if not long_meetings.empty:
            display_df = long_meetings.copy()
            display_df['date'] = display_df['date'].astype(str)
            st.dataframe(
                display_df,
                column_config={
                    "subject": st.column_config.TextColumn("Meeting", width="large"),
                    "duration_minutes": st.column_config.NumberColumn("Duration (min)"),
                    "date": st.column_config.TextColumn("Date"),
                    "organizer": st.column_config.TextColumn("Organizer")
                },
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info("No meetings over 60 minutes found")

    st.divider()

    # Executive Summary Section
    st.markdown('<h2 class="section-header">📝 Executive Summary</h2>', unsafe_allow_html=True)

    # Generate summary
    summary_md = generate_full_summary(
        kpis=kpis,
        patterns=patterns,
        top_meetings=top_meetings,
        top_organizers=top_organizers,
        long_meetings=long_meetings
    )

    # Display summary
    with st.expander("View Full Summary", expanded=True):
        st.markdown(summary_md)

    # Download button
    st.download_button(
        label="📥 Download Summary (Markdown)",
        data=summary_md,
        file_name="calendar_audit_summary.md",
        mime="text/markdown",
        help="Download the executive summary as a Markdown file"
    )

    st.divider()

    # Raw Data Section (collapsible)
    with st.expander("🔍 View Raw Data"):
        st.subheader("Filtered Meeting Data")
        display_cols = ['subject', 'start_datetime', 'end_datetime', 'duration_minutes',
                        'organizer', 'location', 'weekday']
        available_cols = [c for c in display_cols if c in filtered_df.columns]
        st.dataframe(filtered_df[available_cols], use_container_width=True)

        # Download raw data
        csv = filtered_df[available_cols].to_csv(index=False)
        st.download_button(
            label="📥 Download Filtered Data (CSV)",
            data=csv,
            file_name="calendar_audit_data.csv",
            mime="text/csv"
        )


if __name__ == "__main__":
    main()
