/**
 * @fileoverview Syncs a Google Sheet to a Google Calendar.
 * This script creates/updates/deletes 'Busy' events and multi-day 'free' events
 * to make the calendar an exact mirror of the sheet's data.
 *
 * @version 3.1
 * @author wendels
 * @change-log
 * - v3.1 (2025-10-21):
 * - GENERALIZED: Removed all company-specific and personal identifiers from the script.
 * Replaced hardcoded values with configurable placeholders for public sharing.
 */

// --- CONFIGURATION ---
const CONFIG = {
  // TODO: Update this to your desired event title for 'busy' entries.
  busyEventTitle: 'Busy',   // The title for events marked as 'busy'.
  
  calendarId: 'primary',    // Uses the main Google Calendar.
  startColumn: 1,           // Column A
  endColumn: 2,             // Column B
  titleColumn: 3,           // Column C
  availabilityColumn: 5,    // Column E
  headerRows: 1,            // Set to 1 for a header row.
  
  // Configuration for the 'busy' event color.
  // BLUE (ID 9) corresponds to the 'Blueberry' color in the Google Calendar UI.
  busyEventColor: CalendarApp.EventColor.BLUE,

  // **PERFORMANCE OPTIMIZATION**
  syncMonthsPast: 0,        // Set to 0 to prevent changes to past events.
  syncMonthsFuture: 6       // How many months in the future to check.
};
// --- END OF CONFIGURATION ---


/**
 * Adds a custom menu to the spreadsheet UI when it's opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Calendar Sync')
      .addItem('Run Sync Now', 'syncSheetToCalendar')
      .addSeparator()
      .addItem('Cleanup Duplicates (Run Once)', 'cleanupDuplicateEvents')
      .addToUi();
}


/**
 * Sets up a time-driven trigger to run the sync automatically every 15 minutes.
 * Run this function ONCE from the Apps Script editor to set up the automation.
 */
function createTimeDrivenTrigger() {
  // Delete any existing triggers to avoid duplicates.
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'syncSheetToCalendar') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create a new trigger.
  ScriptApp.newTrigger('syncSheetToCalendar')
      .timeBased()
      .everyMinutes(15)
      .create();

  SpreadsheetApp.getUi().alert('Automatic sync has been set up. The script will now run approximately every 15 minutes.');
}


/**
 * Main function to fully sync the sheet and calendar, including deletions.
 */
function syncSheetToCalendar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const allValues = sheet.getRange("A2:E2001").getValues();
  
  const nonEmptyRows = allValues.filter(row => row[CONFIG.startColumn - 1] !== '');
  Logger.log(`Fetched range A2:E2001. Processing ${nonEmptyRows.length} non-empty rows.`);

  if (nonEmptyRows.length === 0) {
    Logger.log('No data to process.');
    return;
  }

  // Define the time window for the sync.
  const now = new Date();
  const syncStartDate = new Date(now);
  syncStartDate.setHours(0, 0, 0, 0); // Start sync window from the beginning of today.
  
  const syncEndDate = new Date(now.getFullYear(), now.getMonth() + CONFIG.syncMonthsFuture + 1, 0);
  Logger.log(`Syncing events between ${syncStartDate.toDateString()} and ${syncEndDate.toDateString()}`);
  Logger.log(`Events starting before ${now.toLocaleString()} will not be deleted.`);


  // 1. Build a "source of truth" Set from the sheet for what SHOULD exist.
  const sheetEventKeys = new Set();
  const oneDayMs = 24 * 60 * 60 * 1000;
  
  const data = nonEmptyRows.filter(row => {
    const d = new Date(row[CONFIG.startColumn - 1]);
    return d >= syncStartDate && d <= syncEndDate;
  });

  for (const row of data) {
    try {
        const start = new Date(row[CONFIG.startColumn - 1]);
        const end = new Date(row[CONFIG.endColumn - 1]);
        const title = row[CONFIG.titleColumn - 1].toString().trim();
        const availability = row[CONFIG.availabilityColumn - 1].toString().trim();

        if (isNaN(start.getTime()) || isNaN(end.getTime())) continue;
        // Example titles to ignore. You can customize these.
        if (title.toLowerCase() === 'pto' || title.toLowerCase() === '[res]' || title.toLowerCase() === '[dev]') continue;

        let eventTitle;
        if (availability.toLowerCase() === 'busy') {
            eventTitle = CONFIG.busyEventTitle;
            const key = `${start.toISOString()}|${end.toISOString()}|${eventTitle}`;
            sheetEventKeys.add(key);
        } else if (availability.toLowerCase() === 'free') {
            const durationMs = end.getTime() - start.getTime();
            if (durationMs >= oneDayMs && title.length <= 5) {
                eventTitle = title;
                const calendarEndDate = new Date(end.getTime());
                calendarEndDate.setDate(calendarEndDate.getDate() + 1);
                
                const key = `${start.toISOString()}|${calendarEndDate.toISOString()}|${eventTitle}`;
                sheetEventKeys.add(key);
            }
        }
    } catch (e) {
        Logger.log(`Error parsing a row. Skipping. Details: ${e.message}`);
    }
  }
  Logger.log(`Found ${sheetEventKeys.size} managed events that should exist based on the sheet.`);

  // 2. Fetch all events from the calendar in the time window.
  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);
  const existingEvents = calendar.getEvents(syncStartDate, syncEndDate);
  Logger.log(`Found ${existingEvents.length} total existing events on the calendar.`);

  let eventsDeleted = 0;
  let eventsCreated = 0;
  let eventsUpdated = 0;
  const calendarEventKeys = new Set();

  // 3. First pass: Identify, update, and delete obsolete managed events.
  for (const event of existingEvents) {
    const eventTitle = event.getTitle();
    const isBusyEvent = eventTitle === CONFIG.busyEventTitle;
    const isManagedFreeEvent = (event.getEndTime().getTime() - event.getStartTime().getTime()) >= oneDayMs && event.getTransparency() === 'TRANSPARENT';

    if (isBusyEvent || isManagedFreeEvent) {
        const key = `${event.getStartTime().toISOString()}|${event.getEndTime().toISOString()}|${eventTitle}`;
        calendarEventKeys.add(key);

        if (!sheetEventKeys.has(key)) {
            if (event.getStartTime() > now) {
                event.deleteEvent();
                Logger.log(`Deleted obsolete event: '${eventTitle}' from ${event.getStartTime()}`);
                eventsDeleted++;
            } else {
                Logger.log(`Skipping deletion of past event: '${eventTitle}' at ${event.getStartTime()}`);
            }
        } else {
            if (isBusyEvent && event.getColor() !== CONFIG.busyEventColor) {
                event.setColor(CONFIG.busyEventColor);
                Logger.log(`Updated color for event: '${eventTitle}' at ${event.getStartTime()}`);
                eventsUpdated++;
            }
        }
    }
  }

  // 4. Second pass: Create new events that are missing from the calendar.
  for (const key of sheetEventKeys) {
    if (!calendarEventKeys.has(key)) {
      const [startIso, endIso, title] = key.split('|');
      const start = new Date(startIso);
      const end = new Date(endIso);

      const newEvent = calendar.createEvent(title, start, end);
      if (newEvent) {
        if (title === CONFIG.busyEventTitle) {
            try {
                newEvent.setColor(CONFIG.busyEventColor);
            } catch (e) {
                Logger.log(`Could not set color for new event '${title}'. Error: ${e.message}`);
            }
        } else {
           try {
               newEvent.setTransparency('TRANSPARENT');
           } catch (e) {
               Logger.log(`Could not set transparency for event '${title}'. The event was still created. Error: ${e.message}`);
           }
        }
        Logger.log(`Created new event: '${title}' from ${start}`);
        eventsCreated++;
      }
    }
  }

  Logger.log(`Sync complete. Created: ${eventsCreated}, Updated: ${eventsUpdated}, Deleted: ${eventsDeleted}.`);
}

/**
 * A one-time utility function to clean up duplicate events created by a previous
 * bug. This version uses a simpler, more direct method to find and remove duplicates.
 */
function cleanupDuplicateEvents() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const allValues = sheet.getRange("A2:E2001").getValues();
  const nonEmptyRows = allValues.filter(row => row[CONFIG.startColumn - 1] !== '');
  Logger.log(`Cleanup: Processing ${nonEmptyRows.length} rows to identify events to check.`);

  // 1. Get a set of unique titles for multi-day "free" events from the sheet.
  const oneDayMs = 24 * 60 * 60 * 1000;
  const titlesToCheck = new Set();
  for (const row of nonEmptyRows) {
    try {
      const start = new Date(row[CONFIG.startColumn - 1]);
      const end = new Date(row[CONFIG.endColumn - 1]);
      const title = row[CONFIG.titleColumn - 1].toString().trim();
      const availability = row[CONFIG.availabilityColumn - 1].toString().trim();

      if (isNaN(start.getTime()) || isNaN(end.getTime())) continue;
      if (availability.toLowerCase() !== 'free') continue;
      
      const durationMs = end.getTime() - start.getTime();
      if (durationMs >= oneDayMs && title.length <= 5) {
        titlesToCheck.add(title);
      }
    } catch (e) { 
      Logger.log(`Cleanup: Error parsing a row. Skipping. Details: ${e.message}`);
    }
  }
  Logger.log(`Cleanup: Will check for duplicates with titles: ${[...titlesToCheck].join(', ')}`);

  // 2. Search for events and clean them up.
  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);
  const now = new Date();
  // Define a broad search window to catch all potential duplicates.
  const searchStartDate = new Date(now.getFullYear(), now.getMonth() - 3, 1);
  const searchEndDate = new Date(now.getFullYear(), now.getMonth() + CONFIG.syncMonthsFuture + 1, 0);
  let deletedCount = 0;

  // 3. For each title, group all matching calendar events by start date and remove extras.
  for (const title of titlesToCheck) {
    const events = calendar.getEvents(searchStartDate, searchEndDate, { search: title });
    
    // Group events by their start date (ignoring time part).
    const eventsByStartDate = new Map();
    for (const event of events) {
      // SIMPLIFIED LOGIC: Only check for an exact title match. This is more aggressive and
      // will catch all duplicates regardless of their other properties (like end date or transparency).
      if (event.getTitle() === title) {
          const startDateKey = event.getStartTime().toISOString().split('T')[0]; // YYYY-MM-DD
          if (!eventsByStartDate.has(startDateKey)) {
            eventsByStartDate.set(startDateKey, []);
          }
          eventsByStartDate.get(startDateKey).push(event);
      }
    }
    
    // 4. Iterate through the grouped events and delete duplicates.
    for (const [dateKey, eventGroup] of eventsByStartDate.entries()) {
      if (eventGroup.length > 1) {
        Logger.log(`Cleanup: Found ${eventGroup.length} duplicates for '${title}' starting on ${dateKey}. Deleting extras.`);
        
        // Keep the first one, delete the rest.
        for (let i = 1; i < eventGroup.length; i++) {
          eventGroup[i].deleteEvent();
          deletedCount++;
        }
      }
    }
  }

  Logger.log(`Cleanup complete. Deleted ${deletedCount} duplicate events.`);
  ui.alert(`Cleanup complete. Deleted ${deletedCount} duplicate events.`);
}


