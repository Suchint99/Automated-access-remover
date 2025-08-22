Google Drive/Sheets Access Manager (Dummy version cause of NDA project agreement)
This is a sanitized demo of a Node.js automation tool that:

What to do:-
Reads data from a Google Spreadsheet.
Flags and highlights rows based on business logic.
Removes Google Drive access for qualifying emails.
Optionally adds a checkbox in a designated column.

What It Does:-
Connects to Google Sheets using a service account.
Checks for users who haven't been paid in over 60 days and were hired at least 60 days ago.
Highlights rows in red and removes access to a set of Google Drive files for those users.

File Overview:-
File	Purpose
worker-access.js	Main script to process sheet and manage Drive access
package.json / package-lock.json	Dependencies
client_secret.json	OAuth2 client credentials (dummy)
service_account.json	Google service account credentials (dummy)

Note: This repository contains dummy credentials and fake IDs for security reasons.

Setup Instructions:-
To use this in your real project:

Replace client_secret.json and service_account.json with your real credentials.
Update the SPREADSHEET_ID and FILE_IDS in worker-access.js with actual values.
Install dependencies:

bash
Copy
Edit
npm install
Run the script:

bash
Copy
Edit
node worker-access.js
Security Tips
NEVER commit real API keys or private keys to GitHub.


