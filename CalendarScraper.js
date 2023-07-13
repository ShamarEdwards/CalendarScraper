const fs = require('fs');
const { google } = require('googleapis');
const XLSX = require('xlsx');

const API_KEY_PATH = './google_api_key.json';

// Load JSON file
const apiKeyData = fs.readFileSync(API_KEY_PATH);
const { api_key, calendar_id } = JSON.parse(apiKeyData);

// Set up Google Calendar API client
const calendar = google.calendar({ version: 'v3', auth: api_key });

// Set the start date to March 1, 2017
const startDate = new Date('2017-03-01T00:00:00Z');

// Set the end date to March 31, 2023
const endDate = new Date('2023-03-31T23:59:59Z');

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Retrieve events for each year from 2017 to 2023
for (let year = 2017; year <= 2023; year++) {
  // Set the start and end dates for the specific year
  startDate.setFullYear(year);
  endDate.setFullYear(year);

  // Create a new worksheet for the year
  const worksheet = XLSX.utils.aoa_to_sheet([['Email Address', 'Event Description']]);

  // Set to store unique email addresses
  const uniqueEmails = new Set();

  // Retrieve events within the specified time range for the year
  calendar.events.list(
    {
      calendarId: calendar_id,
      timeMin: startDate.toISOString(),
      timeMax: endDate.toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
    },
    (err, res) => {
      if (err) {
        console.error('The Google API request returned an error:', err.message);
        return;
      }

      const events = res.data.items;
      if (events.length) {
        console.log(`Email addresses and event descriptions found in Google Calendar events for ${year}:`);

        events.forEach((event) => {
          if (event.description) {
            const emailRegex = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g; // Regular expression for email address matching
            const matches = event.description.match(emailRegex); // Extract email addresses from event description

            if (matches) {
              matches.forEach((email) => {
                if (!email.endsWith('@kemptonasset.com')) {
                  uniqueEmails.add(email);
                  XLSX.utils.sheet_add_aoa(worksheet, [[email, event.description]], { origin: -1 });
                }
              });
            }
          }
        });

        // Add the worksheet to the workbook with the year as the tab name
        XLSX.utils.book_append_sheet(workbook, worksheet, `${year}`);
      } else {
        console.log(`No events found in the specified time range for ${year}.`);
      }

      // Save the workbook as an Excel file after processing all years
      if (year === 2023) {
        const excelFilename = `prospects_emails_year.xlsx`;
        XLSX.writeFile(workbook, `${__dirname}/temp/${excelFilename}`);
        console.log(`Email addresses and event descriptions saved to ${excelFilename}`);
      }
    }
  );
}
