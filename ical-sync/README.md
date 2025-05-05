# Google Sheet iCal Sync

## Purpose

This Google Apps Script synchronizes event data from multiple iCalendar (`.ics`) URLs into a designated Google Sheet. It fetches events, parses them, adds new events, updates existing ones, and marks events present in the sheet but missing from the iCal feeds as potentially cancelled.

## Features

*   Reads iCal URLs from a configuration sheet.
*   Processes multiple calendars.
*   Creates/Updates a "Bookings" sheet with event details (UID, Property Name, Status, Check-in/Check-out dates, Nights, Last Updated).
*   Translates standard iCal statuses (CONFIRMED, TENTATIVE, CANCELLED) to Traditional Chinese.
*   Identifies and flags events potentially removed from the source iCal feeds.
*   Adds a custom menu item for manual synchronization.

## Setup and Configuration

1.  **Make a Copy:** Create a copy of the Google Sheet containing this script.
2.  **Open Script Editor:** In your copied Sheet, go to `Tools` > `Script editor`.
3.  **Configure Calendars:**
    *   Open the sheet named `日曆連結` (or `ICAL_CONFIG_SHEET_NAME` if you changed it).
    *   Add your calendars:
        *   **Column A (房源名稱):** A unique name for the property/calendar source.
        *   **Column B (iCal 網址):** The full `http` or `https` URL of the iCal feed.
        *   **Column C (啟用):** Set this to `是` for the script to process this URL. Leave blank or set to anything else to disable.
4.  **Check Target Sheet:**
    *   The script will write data to the sheet named `預訂紀錄` (or `BOOKINGS_SHEET_NAME`).
    *   If this sheet doesn't exist, the script will create it with the necessary headers.
5.  **(Optional) Change Sheet Names:**
    *   If you rename the `日曆連結` or `預訂紀錄` sheets, you **must** update the corresponding constants at the top of the `Code.ts` file in the Script Editor:
        ```typescript
        const ICAL_CONFIG_SHEET_NAME = "Your Config Sheet Name";
        const BOOKINGS_SHEET_NAME = "Your Bookings Sheet Name";
        ```
6.  **Authorize:** The first time you run the sync (or open the sheet after copying), Google will ask for authorization to access external services (fetch URLs) and modify your spreadsheet. Grant the necessary permissions.

## Running the Script

*   **Manual Sync:**
    *   Open the Google Sheet.
    *   Use the custom menu: `iCal Sync` > `同步預訂紀錄`.
    *   Alternatively, open the Script Editor (`Tools` > `Script editor`), select the function `syncAllIcalLinks` from the dropdown menu, and click `Run`.
*   **Automatic Sync (Triggers):**
    *   In the Script Editor, click the `Triggers` icon (looks like an alarm clock) in the left sidebar.
    *   Click `+ Add Trigger`.
    *   Configure the trigger:
        *   Choose function to run: `syncAllIcalLinks`
        *   Choose deployment: `Head`
        *   Select event source: `Time-driven`
        *   Select type of time based trigger: (e.g., `Hourly timer`, `Daily timer`)
        *   Set error notification settings.
    *   Click `Save`. You may need to authorize again.

## Sheets
* [Development Sheet](https://docs.google.com/spreadsheets/d/1Ui1R1aQSZY3hZw6c-3dvfeuGS_s5VbkofaRTBsriM90/edit?usp=sharing)
* [Production Sheet](https://docs.google.com/spreadsheets/d/1DQ2dErxDnv9WX-FqR5_MpIFT1Ywz4Ran8u6XqKRwkuk/edit?usp=sharing)

## Customization

*   **Status Translation:** Modify the `translateStatus` function in `Code.ts` if you need different status names or languages.
