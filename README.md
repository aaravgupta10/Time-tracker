# Personal Analytics Dashboard 📊

This repository contains the Google Apps Script code that powers my daily personal productivity analytics engine.

I use this system daily to gain data-driven insights into my time management. It automatically syncs my Google Calendars to a master Spreadsheet, processes the data, and emails me detailed reports on my time allocation across various aspects of my life (Academics, Health, Productivity, etc.).

## 🚀 How It Works

1.  **Data Ingestion:**
    -   The script runs automatically every morning via Google Apps Script Triggers.
    -   It fetches the previous day's events from multiple specific Google Calendars (e.g., "College", "Deep Work", "Health").
    -   It cleans and formats this data (calculating duration, assigning quarters/weeks) and logs it into a "Raw Data" warehouse in Google Sheets.

2.  **Analysis & Visualization:**
    -   The data is processed in the Spreadsheet using pivot tables and custom formulas to categorize time into "Productive" vs. "Unproductive" buckets.
    -   *(Note: The logic for the spreadsheet formulas is contained within the sheet itself, while this repo hosts the automation backend.)*

3.  **Reporting:**
    -   **Daily Reports:** Sent at 7:00 AM with a breakdown of yesterday's metrics.
    -   **Weekly/Monthly/Quarterly/Annual Reports:** Sent automatically at the start of each period to review long-term trends and goal alignment.

## 🛠️ Tech Stack

-   **Google Apps Script (JavaScript):** Serverless backend logic.
-   **Google Calendar API:** For fetching event data.
-   **Google Sheets API:** For data storage and report generation.
-   **Gmail API:** For sending HTML-formatted email reports.

## 🔐 Configuration & Security

To maintain security and separate configuration from code, this script uses **Google Apps Script Properties Service**.

-   **`SHEET_ID`**: The unique ID of the Google Sheet database.
-   **`EMAIL_RECIPIENT`**: The target email address for reports.

These values are stored as script properties (environment variables) and are not hardcoded in the source file.

## ⚙️ Setup Overview

The `CONFIG` object at the top of the script allows for customization of:
-   **Tracked Calendars:** Which specific calendars to pull data from.
-   **Productivity Mapping:** Defining which calendars contribute to "Productive" hours.
-   **Timezone:** Automatically syncs with the script's execution timezone.

## 📈 Usage

This script is deployed as a standalone project attached to my Google account. It operates entirely in the background, requiring no manual input once configured.
