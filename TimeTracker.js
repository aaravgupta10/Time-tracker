/**************************************************************************
 * PERSONAL ANALYTICS DASHBOARD - CONFIGURATION
 * ------------------------------------------------------------------------
 * This is our main settings panel.
 **************************************************************************/
const CONFIG = {
    // 1. Get this from your Google Sheet URL: .../d/[THISIS_THE_ID]/edit
    // SECURE: Fetching ID from Script Properties to avoid hardcoding secrets
    SHEET_ID: PropertiesService.getScriptProperties().getProperty('SHEET_ID'),

    // 2. The email address to send reports to.
    // SECURE: Fetching Email from Script Properties
    EMAIL_RECIPIENT: PropertiesService.getScriptProperties().getProperty('EMAIL_RECIPIENT'),

    // 3. (Optional) An email to CC on reports.
    EMAIL_CC: "",

    // 4. The exact names of the calendars you want to track.
    CALENDARS_TO_TRACK: [
        "College",
        "Morning & personal routine",
        "Productivity & Work",
        "Sleep",
        "Tennis & Health",
        "Unproductive"
    ],

    // 5. Define what counts as "Productive" for your reports.
    PRODUCTIVE_CALENDARS: [
        "College",
        "Productivity & Work",
        "Morning & personal routine",
        "Tennis & Health"
    ],

    // 6. Define what counts as "Unproductive"
    UNPRODUCTIVE_CALENDARS: ["Unproductive"],

    // 7. Your script's timezone (usually correct by default).
    TIMEZONE: Session.getScriptTimeZone()
};

/**************************************************************************
 * AESTHETICS CONFIGURATION
 **************************************************************************/
const AESTHETICS = {
    headerBg: "#4a86e8", // A nice blue for headers
    headerFont: "#ffffff", // White text
    sheetTabColor: "#4a86e8", // Color for the sheet tabs
    fontFamily: "Arial"
};


/**************************************************************************
 * 1. MENU & SETUP FUNCTIONS
 **************************************************************************/

/**
 * Runs when the spreadsheet is opened. Creates a custom menu.
 * Updated to include the "Day Before Yesterday" button.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Personal Analytics')
        .addItem('1. Run Initial Setup', 'runInitialSetup')
        .addSeparator()
        .addItem('Fetch Yesterday\'s Events (Test)', 'fetchAndLogEvents')
        .addItem('Fetch Day Before Yesterday (Manual)', 'fetchDayBeforeYesterdayEvents') // <-- NEW BUTTON
        .addSeparator()
        .addItem('Send Daily Report (Test)', 'sendDailyReport')
        .addItem('Send Weekly Report (Test)', 'sendWeeklyReport')
        .addItem('Send Monthly Report (Test)', 'sendMonthlyReport')
        .addItem('Send Quarterly Report (Test)', 'sendQuarterlyReport')
        .addItem('Send Annual Report (Test)', 'sendAnnualReport')
        .addSeparator()
        .addItem('List All My Calendar Names (Debug)', 'listMyCalendarNames')
        .addToUi();
}

/**
 * A helper function to list all your calendar names for debugging.
 */
function listMyCalendarNames() {
    const allCalendars = CalendarApp.getAllCalendars();
    const names = allCalendars.map(cal => cal.getName());
    Logger.log("--- YOUR CALENDAR NAMES ---");
    Logger.log(names.join("\n"));
    Logger.log("---------------------------");
    SpreadsheetApp.getUi().alert("Your calendar names have been printed to the Executions log. Please check the log.");
}

/**
 * Main setup function. We will run this to build the sheet.
 */
function runInitialSetup() {
    Logger.log("Starting initial setup...");
    setupSheetStructure();
    createTriggers();
    Logger.log("Setup complete. Sheets are created and triggers are set.");
}

/**
 * Creates all the required tabs and headers in the spreadsheet.
 * This is the FINAL, CORRECTED version.
 */
function setupSheetStructure() {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

    // Define all the tabs we need
    const requiredTabs = {
        "Raw Data": ["Date", "Event Title", "Start Time", "End Time", "Duration (Decimal)", "Calendar", "Week-Year", "Month-Year", "Quarter-Year", "Year"],
        "Config": ["Productive Calendars", "Unproductive Calendars", "All Tracked Calendars", "Sheet ID"],
        "Analysis_Engine (Pivot)": [],
        "Dashboard": [],
        "Reports": []
    };

    Object.keys(requiredTabs).forEach(tabName => {
        let sheet = ss.getSheetByName(tabName);
        if (!sheet) {
            sheet = ss.insertSheet(tabName);
            Logger.log(`Created tab: ${tabName}`);
            sheet.setFrozenRows(1); // Only freeze rows for new sheets
        }

        // --- THIS IS THE BUG FIX ---
        // This code is now OUTSIDE the 'if' block.
        // It will run EVERY time, ensuring headers are always correct.
        if (requiredTabs[tabName].length > 0) {
            // Get the full header range
            const headerRange = sheet.getRange(1, 1, 1, requiredTabs[tabName].length);
            // Set the new, correct headers
            headerRange.setValues([requiredTabs[tabName]]).setFontWeight("bold");
        }
        // --- END OF BUG FIX ---

        // Apply aesthetics
        sheet.setTabColor(AESTHETICS.sheetTabColor);
        applyHeaderFormatting(sheet);
    });

    // Populate the Config tab
    const configSheet = ss.getSheetByName("Config");
    configSheet.getRange("A2:A").clearContent();
    configSheet.getRange("B2:B").clearContent();
    configSheet.getRange("C2:C").clearContent();

    configSheet.getRange(2, 1, CONFIG.PRODUCTIVE_CALENDARS.length, 1).setValues(CONFIG.PRODUCTIVE_CALENDARS.map(c => [c]));
    configSheet.getRange(2, 2, CONFIG.UNPRODUCTIVE_CALENDARS.length, 1).setValues(CONFIG.UNPRODUCTIVE_CALENDARS.map(c => [c]));
    configSheet.getRange(2, 3, CONFIG.CALENDARS_TO_TRACK.length, 1).setValues(CONFIG.CALENDARS_TO_TRACK.map(c => [c]));

    // Add the Sheet ID to the Config tab so formulas can use it
    configSheet.getRange("D2").setValue(CONFIG.SHEET_ID)
        .setFontFamily(AESTHETICS.fontFamily);
    applyHeaderFormatting(configSheet); // Re-apply formatting

    // Format the Raw Data sheet columns
    applyDataFormatting(ss.getSheetByName("Raw Data"));

    Logger.log("Populated Config tab and applied formatting.");
}


/**************************************************************************
 * 2. DATA FETCHING & FORMATTING
 **************************************************************************/

/**
 * Fetches events from *yesterday* for all tracked calendars and logs to "Raw Data".
 * This is the FINAL version.
 */
function fetchAndLogEvents() {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const dataSheet = ss.getSheetByName("Raw Data");

    // Set time range for *all of yesterday*
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0); // Start of yesterday

    const endOfYesterday = new Date(yesterday);
    endOfYesterday.setHours(23, 59, 59, 999); // End of yesterday

    // --- NEW: Generate all date formats we need ---
    const year = Utilities.formatDate(yesterday, CONFIG.TIMEZONE, "yyyy");
    const yearWeek = Utilities.formatDate(yesterday, CONFIG.TIMEZONE, "yyyy-'W'WW"); // e.g., "2025-W44"
    const yearMonth = Utilities.formatDate(yesterday, CONFIG.TIMEZONE, "yyyy-MM"); // e.g., "2025-10"
    const quarter = "Q" + Math.floor((yesterday.getMonth() + 3) / 3); // e.g., "Q4"
    const yearQuarter = year + "-" + quarter; // e.g., "2025-Q4"
    // ---

    Logger.log(`Fetching events for ${yesterday} (Week: ${yearWeek}, Month: ${yearMonth})`);

    // Get all calendars and filter for the ones we want
    const allCalendars = CalendarApp.getAllCalendars();
    const targetCalendars = allCalendars.filter(cal => CONFIG.CALENDARS_TO_TRACK.includes(cal.getName()));

    if (targetCalendars.length === 0) {
        Logger.log("Error: No target calendars found. Check calendar names in CONFIG.CALENDARS_TO_TRACK.");
        return; // Stop the function
    }
    Logger.log(`Found ${targetCalendars.length} calendars to track.`);

    let newEventsData = [];

    // Loop through each target calendar and get events
    targetCalendars.forEach(calendar => {
        const calName = calendar.getName();
        try {
            const events = calendar.getEvents(yesterday, endOfYesterday);

            events.forEach(event => {
                const startTime = event.getStartTime();
                const endTime = event.getEndTime();
                // Duration in decimal hours for calculations
                const durationDecimal = (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60);

                if (event.isAllDayEvent()) return; // Skip all-day events

                // --- NEW: Add all new date columns to our data log ---
                newEventsData.push([
                    yesterday, // Date
                    event.getTitle(),
                    startTime,
                    endTime,
                    durationDecimal,
                    calName,
                    yearWeek,
                    yearMonth,
                    yearQuarter,
                    year
                ]);
                // ---
            });
        } catch (e) {
            Logger.log(`Error fetching calendar ${calName}: ${e}`);
        }
    });

    if (newEventsData.length > 0) {
        // Append all new data in one go
        dataSheet.getRange(dataSheet.getLastRow() + 1, 1, newEventsData.length, newEventsData[0].length)
            .setValues(newEventsData);
        Logger.log(`Successfully fetched ${newEventsData.length} new events from yesterday.`);

        // Auto-format the data we just added
        applyDataFormatting(dataSheet);

    } else {
        Logger.log(`No new events found for yesterday.`);
    }
}

/**
 * Applies aesthetic formatting to the headers of a sheet.
 */
function applyHeaderFormatting(sheet) {
    if (sheet.getFrozenRows() === 0) return; // Don't format if no header

    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground(AESTHETICS.headerBg)
        .setFontColor(AESTHETICS.headerFont)
        .setFontFamily(AESTHETICS.fontFamily)
        .setFontWeight("bold");
}

/**
 * Applies formatting to the data area of the Raw Data sheet.
 */
function applyDataFormatting(sheet) {
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());

    // Set base font for all data
    dataRange.setFontFamily(AESTHETICS.fontFamily);

    // Set specific number formats for readability
    sheet.getRange("A2:A").setNumberFormat("yyyy-mm-dd"); // Date
    sheet.getRange("C2:C").setNumberFormat("h:mm am/pm"); // Start Time
    sheet.getRange("D2:D").setNumberFormat("h:mm am/pm"); // End Time
    sheet.getRange("E2:E").setNumberFormat("0.00"); // Duration (Decimal)

    // Apply alternating row colors for readability
    try {
        const banding = dataRange.getBandings()[0];
        if (banding) {
            banding.remove();
        }
        dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    } catch (e) {
        Logger.log("Error applying banding: " + e);
    }
}
/**
 * Fetches events from *2 days ago* for all tracked calendars and logs to "Raw Data".
 * Useful if the script failed to run or you need to backfill data manually.
 */
function fetchDayBeforeYesterdayEvents() {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const dataSheet = ss.getSheetByName("Raw Data");

    // Set time range for *day before yesterday* (Today - 2 days)
    const targetDate = new Date();
    targetDate.setDate(targetDate.getDate() - 2); // <-- The Key Change
    targetDate.setHours(0, 0, 0, 0); // Start of day

    const endOfTargetDate = new Date(targetDate);
    endOfTargetDate.setHours(23, 59, 59, 999); // End of day

    // --- Generate date formats ---
    const year = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, "yyyy");
    const yearWeek = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, "yyyy-'W'WW");
    const yearMonth = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, "yyyy-MM");
    const quarter = "Q" + Math.floor((targetDate.getMonth() + 3) / 3);
    const yearQuarter = year + "-" + quarter;
    // ---

    Logger.log(`Fetching events for ${targetDate} (Week: ${yearWeek})`);

    // Get calendars
    const allCalendars = CalendarApp.getAllCalendars();
    const targetCalendars = allCalendars.filter(cal => CONFIG.CALENDARS_TO_TRACK.includes(cal.getName()));

    if (targetCalendars.length === 0) {
        Logger.log("Error: No target calendars found.");
        return;
    }

    let newEventsData = [];

    // Loop through calendars
    targetCalendars.forEach(calendar => {
        const calName = calendar.getName();
        try {
            const events = calendar.getEvents(targetDate, endOfTargetDate);

            events.forEach(event => {
                const startTime = event.getStartTime();
                const endTime = event.getEndTime();
                const durationDecimal = (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60);

                if (event.isAllDayEvent()) return;

                newEventsData.push([
                    targetDate, // Date
                    event.getTitle(),
                    startTime,
                    endTime,
                    durationDecimal,
                    calName,
                    yearWeek,
                    yearMonth,
                    yearQuarter,
                    year
                ]);
                // ---
            });
        } catch (e) {
            Logger.log(`Error fetching calendar ${calName}: ${e}`);
        }
    });

    if (newEventsData.length > 0) {
        dataSheet.getRange(dataSheet.getLastRow() + 1, 1, newEventsData.length, newEventsData[0].length)
            .setValues(newEventsData);
        Logger.log(`Successfully fetched ${newEventsData.length} events from day before yesterday.`);
        applyDataFormatting(dataSheet);
        SpreadsheetApp.getUi().alert(`Success: Fetched ${newEventsData.length} events from 2 days ago.`);
    } else {
        Logger.log(`No events found for day before yesterday.`);
        SpreadsheetApp.getUi().alert("No events found for 2 days ago.");
    }
}


/**************************************************************************
 * 3. AUTOMATION & TRIGGERS
 **************************************************************************/

/**
 * Creates all the time-driven triggers for automation.
 * This is the FINAL, CORRECTED version with 4 AM fetch time.
 */
function createTriggers() {
    deleteAllTriggers(); // Clear old triggers first

    Logger.log("Creating new triggers...");

    // --- THIS IS THE CHANGE ---
    // 1. Fetch data every morning for the previous day.
    ScriptApp.newTrigger('fetchAndLogEvents')
        .timeBased()
        .atHour(4) // <-- Was 1, is now 4. Runs at 4:00 AM
        .everyDays(1)
        .create();
    Logger.log("Created trigger: fetchAndLogEvents (Daily @ 4 AM)");
    // --- END OF CHANGE ---

    // 2. Send Daily Report (reports on *yesterday*)
    ScriptApp.newTrigger('sendDailyReport')
        .timeBased()
        .atHour(7) // Runs at 7:00 AM
        .everyDays(1)
        .create();
    Logger.log("Created trigger: sendDailyReport (Daily @ 7 AM)");

    // 3. Send Weekly Report
    ScriptApp.newTrigger('sendWeeklyReport')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(8) // Runs Monday at 8:00 AM
        .create();
    Logger.log("Created trigger: sendWeeklyReport (Monday @ 8 AM)");

    // 4. Send Monthly Report
    ScriptApp.newTrigger('sendMonthlyReport')
        .timeBased()
        .onMonthDay(1) // Runs on the 1st of every month
        .atHour(8)
        .create();
    Logger.log("Created trigger: sendMonthlyReport (1st @ 8 AM)");

    // 5. Send Quarterly Report
    ScriptApp.newTrigger('sendQuarterlyReport')
        .timeBased()
        .onMonthDay(1)
        .atHour(9)
        .create();
    Logger.log("Created trigger: sendQuarterlyReport (1st @ 9 AM)");

    // 6. Send Annual Report
    ScriptApp.newTrigger('sendAnnualReport')
        .timeBased()
        .onMonthDay(1)
        .atHour(9)
        .create();
    Logger.log("Created trigger: sendAnnualReport (1st @ 9 AM)");

    Logger.log("All triggers created.");
}

/**
 * Helper function to delete all existing triggers.
 * This prevents creating duplicate triggers every time setup is run.
 */
function deleteAllTriggers() {
    const allTriggers = ScriptApp.getProjectTriggers();
    for (const trigger of allTriggers) {
        ScriptApp.deleteTrigger(trigger);
    }
    Logger.log(`Deleted ${allTriggers.length} old triggers.`);
}


/**************************************************************************
 * 4. EMAIL REPORTING FUNCTIONS
 **************************************************************************/

/**
 * These are the wrapper functions our triggers will call.
 * This is the FINAL, CORRECTED version.
 */
function sendDailyReport() {
    generateAndSendEmail('daily');
}
function sendWeeklyReport() {
    generateAndSendEmail('weekly');
}
function sendMonthlyReport() {
    generateAndSendEmail('monthly');
}

function sendQuarterlyReport() {
    const currentMonth = new Date().getMonth(); // 0 = Jan, 1 = Feb, etc.
    // Only run in Jan (0), Apr (3), Jul (6), Oct (9)
    if (currentMonth % 3 === 0) {
        Logger.log("It's the first month of a quarter. Sending Quarterly Report.");
        generateAndSendEmail('quarterly');
    } else {
        Logger.log("Not the first month of a quarter. Skipping Quarterly Report.");
    }
}

function sendAnnualReport() {
    const currentMonth = new Date().getMonth(); // 0 = Jan
    // Only run in Jan (0)
    if (currentMonth === 0) {
        Logger.log("It's January. Sending Annual Report.");
        generateAndSendEmail('annual');
    } else {
        Logger.log("Not January. Skipping Annual Report.");
    }
}

/**
 * The main email generation engine.
 * It reads pre-compiled reports from the 'Reports' tab and emails them.
 * This is the FINAL, CORRECTED version.
 */
/**
 * The main email generation engine.
 * This is the FINAL, CORRECTED version with the new layout.
 */
function generateAndSendEmail(frequency) {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        const reportSheet = ss.getSheetByName("Reports");

        // Define where to find the subject and body in the 'Reports' tab
        // --- THIS IS THE FINAL LAYOUT FIX ---
        // This gives us 2 columns for data + 2 columns for content for each report.
        const reportCells = {
            'daily': { subject: 'F4', body: 'F5' },     // Daily Content (Cols E-F)
            'weekly': { subject: 'J4', body: 'J5' },    // Weekly Content (Cols I-J)
            'monthly': { subject: 'N4', body: 'N5' },   // Monthly Content (Cols M-N)
            'quarterly': { subject: 'R4', body: 'R5' }, // Quarterly Content (Cols Q-R)
            'annual': { subject: 'V4', body: 'V5' }     // Annual Content (Cols U-V)
        };
        // --- END OF FIX ---

        const cells = reportCells[frequency];
        if (!cells) {
            Logger.log(`Invalid report frequency: ${frequency}`);
            return;
        }

        // Recalculate the sheet to ensure all formulas are up to date
        SpreadsheetApp.flush();

        const subject = reportSheet.getRange(cells.subject).getDisplayValue();

        // This is the bug fix from before (it's still here)
        const body = reportSheet.getRange(cells.body).getDisplayValue();
        // 

        // Check if the report body is empty or a placeholder
        if (!body || body.length < 20 || body.startsWith("=")) {
            Logger.log(`Report body for ${frequency} is empty or invalid. Skipping email.`);
            return;
        }

        GmailApp.sendEmail(CONFIG.EMAIL_RECIPIENT, subject, "", {
            htmlBody: body,
            cc: CONFIG.EMAIL_CC
        });

        Logger.log(`Successfully sent ${frequency} report to ${CONFIG.EMAIL_RECIPIENT}.`);

    } catch (e) {
        Logger.log(`Failed to send ${frequency} email: ${e}`);
    }
}