# Privacy Policy for Tiller Tools

**Last updated:** March 24, 2026

## Overview

Tiller Tools is a Google Sheets add-on that allows users to import Amazon order history data into their personal spreadsheets. Processing occurs in your Google account and browser; the developer does not operate servers that receive your spreadsheet or Amazon export contents, except that your browser loads the JSZip script from a public CDN as described below.

## Information We Collect

Tiller Tools does not collect, store, or transmit personal data outside of the user’s Google account.

All data processed by the add-on remains within:

- The user’s Google Sheets
- The user’s browser (during file upload and processing)

The Amazon import sidebar loads the open-source **JSZip** library from a public CDN (cdnjs) in your browser only to unpack a ZIP file you select; that request goes to the CDN operator, not to servers run by the developer, and your files are not sent to the developer’s systems for that step.

## How Data Is Used

The add-on processes user-provided data solely for the purpose of:

- Parsing Amazon order history files uploaded by the user
- Formatting transaction data for compatibility with Tiller spreadsheets
- Writing transaction data into the user’s Google Sheets

No data is used for analytics, advertising, or tracking.

## Data Storage

Tiller Tools does not maintain any external databases or servers.

- No user data is stored outside of Google Sheets
- No data is retained by the developer
- All processing occurs within Google Apps Script and the user’s browser

## Data Sharing

The developer does not receive your spreadsheet or Amazon export contents for analytics, advertising, or resale.

Aside from the browser loading **JSZip** from the CDN as described above, the add-on does not call third-party **APIs** to send your data to other services; CSV payloads are processed with **Google Apps Script** within Google’s infrastructure to update **your** spreadsheet.

## Permissions

The add-on requests access to:

- **Google Sheets** — to read and write transaction data in the user’s spreadsheet
- **Google Apps Script UI** — to display the sidebar interface

These permissions are used solely to provide the functionality of the add-on.

## Security

All data processing occurs within Google’s secure environment, including:

- Google Sheets
- Google Apps Script runtime
- The user’s local browser

Aside from your browser fetching the JSZip script from the CDN, no additional data transmission channels are used by the developer.

## User Control

Users have full control over their data:

- All imported data resides in their own Google Sheets
- Users may edit or delete data at any time
- Users may uninstall the add-on at any time

## Changes to This Policy

This privacy policy may be updated from time to time. Updates will be reflected by the “Last updated” date at the top of this document.

## Contact

If you have any questions about this Privacy Policy, you can contact:

- **Developer:** Tiller Tools  
- **Email:** [tillertools.app@gmail.com](mailto:tillertools.app@gmail.com)

See also: [Terms of Service](TERMS.md).
