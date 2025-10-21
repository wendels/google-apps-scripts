/**
 * @OnlyCurrentDoc
 *
 * Fetches events from a user's Google Calendar and populates them into the
 * active Google Sheet. Designed to be a reusable and configurable template.
 *
 * @author wendels
 * @version 1.0.0
 *
 * --- SETUP INSTRUCTIONS ---
 * 1. Open the Apps Script editor (Extensions > Apps Script).
 * 2. Configure the variables in the "Core Configuration" section below.
 * 3. Run the `getCalendarEvents` function once manually to authorize permissions.
 * 4. Set up a time-driven trigger to run the `getCalendarEvents` function
 * automatically (e.g., every 10 or 15 minutes).
 *
 * --- FEATURES ---
 * - Fetches events from any specified Google Calendar.
 * - Configurable look-ahead period and record limit.
 * - Intelligent trigger schedule runs more often during the day.
 * - Clears and rewrites data to the top of the active sheet.
 */
function getCalendarEvents() {
  // --- START: Core Configuration ---
  // These are the main variables you might want to adjust.

  // The number of days into the future to fetch calendar events.
  const DAYS_TO_FETCH = 90;

  // The maximum number of event records to write to the sheet.
  // This prevents the sheet from becoming too large and slow.
  const MAX_RECORDS_TO_WRITE = 2000;

  // The ID of the calendar to fetch events from.
  // 'default' uses the primary calendar of the user running the script.
  // To use a different calendar, replace 'default' with the calendar's ID
  // (e.g., 'your.email@example.com' or 'xxxxxxxx@group.calendar.google.com').
  const CALENDAR_ID = 'default';

  // --- END: Core Configuration ---


  const now = new Date();
  const properties = PropertiesService.getScriptProperties();
  const lastRunTimestamp = parseInt(properties.getProperty('lastRunTimestamp') || '0');

  // --- Dynamic Execution Schedule ---
  // This logic determines if the script should run based on the time of day.
  const TEN_MINUTES = 10 * 60 * 1000;
  const TWENTY_MINUTES = 20 * 60 * 1000;
  const timeSinceLastRun = now.getTime() - lastRunTimestamp;
  const currentHour = now.getHours();
  let shouldRun = false;

  // Daytime schedule: 9:00 AM to 5:59 PM (hour is 9 to 17). Runs every 10+ minutes.
  if (currentHour >= 9 && currentHour < 18) {
    if (timeSinceLastRun >= TEN_MINUTES) {
      shouldRun = true;
    }
  }
  // Nighttime schedule: Before 9:00 AM or 6:00 PM onwards. Runs every 20+ minutes.
  else {
    if (timeSinceLastRun >= TWENTY_MINUTES) {
      shouldRun = true;
    }
  }

  if (!shouldRun) {
    console.log(`Skipping execution. Last run was ${Math.floor(timeSinceLastRun / 60000)} minutes ago.`);
    return; // Exit if not enough time has passed.
  }

  console.log("Starting script execution...");
  const scriptStartTime = new Date();

  const sheet = SpreadsheetApp.getActiveSheet();
  const NUM_COLUMNS = 5; // Corresponds to the number of headers below.

  // Clear the content of the target range to remove old data.
  sheet.getRange(1, 1, MAX_RECORDS_TO_WRITE + 1, NUM_COLUMNS).clearContent();

  // Set the time range for fetching events.
  const futureDate = new Date();
  futureDate.setDate(now.getDate() + DAYS_TO_FETCH);

  let calendar;
  let allEvents = [];

  try {
    console.log(`Fetching events from calendar: ${CALENDAR_ID}`);
    const fetchStartTime = new Date();
    calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      throw new Error(`Calendar with ID "${CALENDAR_ID}" not found or accessible.`);
    }
    allEvents = calendar.getEvents(now, futureDate);
    const fetchEndTime = new Date();
    console.log(`Fetched ${allEvents.length} events in ${(fetchEndTime - fetchStartTime) / 1000}s.`);
  } catch (e) {
    console.error(`Error fetching events: ${e.message}`);
    sheet.getRange("A1").setValue(`Error: Could not retrieve events. ${e.message}`);
    return; // Exit script on error.
  }

  // Sort events chronologically.
  allEvents.sort((a, b) => a.getStartTime() - b.getStartTime());

  // Truncate the event list if it exceeds the max record limit.
  if (allEvents.length > MAX_RECORDS_TO_WRITE) {
    console.log(`Truncating event list from ${allEvents.length} to ${MAX_RECORDS_TO_WRITE}.`);
    allEvents = allEvents.slice(0, MAX_RECORDS_TO_WRITE);
  }

  // Define the header row for the sheet.
  const header = [
    "Event Start Time",
    "Event End Time",
    "Event Title",
    "Visibility",
    "My Status"
  ];

  // Map the event data to a 2D array for efficient writing to the sheet.
  const eventDetails = allEvents.map(event => {
    const myStatus = String(event.getTransparency()) === 'TRANSPARENT' ? "Free" : "Busy";
    return [
      event.getStartTime(),
      event.getEndTime(),
      event.getTitle(),
      event.getVisibility().toString(),
      myStatus
    ];
  });

  const outputData = [header, ...eventDetails];

  console.log(`Writing ${outputData.length} rows to the sheet...`);
  if (outputData.length > 1) {
    sheet.getRange(1, 1, outputData.length, header.length).setValues(outputData);
  } else {
    // Write only the header if no events were found.
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.getRange(2, 1).setValue(`No events found in the next ${DAYS_TO_FETCH} days.`);
  }
  console.log("Finished writing to sheet.");

  // Store the timestamp of this successful run.
  properties.setProperty('lastRunTimestamp', now.getTime().toString());

  const scriptEndTime = new Date();
  console.log(`Script finished successfully. Total execution time: ${(scriptEndTime - scriptStartTime) / 1000} seconds.`);
}
