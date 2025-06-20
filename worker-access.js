const { google } = require('googleapis');
const { GoogleAuth } = require('google-auth-library');
const moment = require('moment');

const SERVICE_ACCOUNT_PATH = './service_account.json';
const SPREADSHEET_ID = 'your-spreadsheet-id';

// Configuration
const FILE_IDS = [
  'file-id-1',
  'file-id-2',
  '1OnWBQ6XOL9GiW_U4krLrL3PjZ_PuRa0Sxxg4ZjStbMA',
  '15cbT3CqK_eB0aK5IycJp0Y0OquD3ZFpF4B5jqJPTP_M',
  '1Kuku8DqUYUQW3r2iOmGh77FuPnTWIqOJZFpuij0AjYY',
  '1TDFBd0irt5SuDn2TyQq0vfiyOoEC75uQDil9kA6HccM',
  '1JzVZ2iphrQpHZIgyBkmZh3cBg-V7II65yldZj5ldEg8',
  '1o6H9E_PKg4DESEfjPh2wyuARdktAabGVfqzt8Jf_3gI',
  '1mKiREM9IFMI8p23C71zQuSFmRGRvpNS0'
];
const DRY_RUN = false;
const START_ROW = 169; // starting destination row

// Initialize auth
const auth = new GoogleAuth({
  keyFile: SERVICE_ACCOUNT_PATH,
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
  ]
});

// Existing removeDriveAccess function
async function removeDriveAccess(drive, email, fileIds) {
  for (const fileId of fileIds) {
    try {
      console.log(`Checking permissions for ${email} on file ${fileId}...`);
      
      const permissions = await drive.permissions.list({
        fileId,
        fields: 'permissions(id,emailAddress)'
      });

      const permission = permissions.data.permissions.find(
        p => p.emailAddress === email
      );

      if (permission) {
        if (!DRY_RUN) {
          await drive.permissions.delete({
            fileId,
            permissionId: permission.id
          });
          console.log(`✅ Removed access for ${email} from file ${fileId}`);
        } else {
          console.log(`🔄 Would remove access for ${email} from file ${fileId}`);
        }
      } else {
        console.log(`ℹ️ No access found for ${email} on file ${fileId}`);
      }
    } catch (error) {
      console.error(`❌ Error processing ${fileId}:`, error.message);
    }
  }
}

async function main() {
  try {
    console.log('🚀 Starting worker access management process...');
    const sheets = google.sheets({ version: 'v4', auth });
    const drive = google.drive({ version: 'v3', auth });

    // Get spreadsheet data
    const response = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      includeGridData: true
    });

    const sheet = response.data.sheets[0];
    const sheetId = sheet.properties.sheetId;
    const data = sheet.data[0].rowData;
    
    // Find required columns (within A–M)
    const lastPaidColIndex = 11; // Column L
    const emailColIndex = data[0].values.findIndex(
      cell => cell.formattedValue?.toLowerCase().includes('email')
    );
    const hireDateColIndex = 9;  // Column J
    const columnMColIndex = 12;  // Column M

    if (emailColIndex === -1 || emailColIndex >= 13) {
      throw new Error('Email column not found within columns A–M');
    }

    const sixtyDaysAgo = moment().subtract(60, 'days');

    const updateRequests = [];
    const workersToProcess = [];
    const confirmedRowsToMove = []; // will store row indices of access-denied workers
    
    console.log(`🔍 Processing ${data.length - 1} rows...`);
    
    // Step 1: Process qualifying workers
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (!row.values || row.values.length < 13) {
        console.log(`⏩ Skipping row ${i} - insufficient columns`);
        continue;
      }

      const email = row.values[emailColIndex]?.formattedValue?.trim();
      const lastPaidDate = row.values[lastPaidColIndex]?.formattedValue?.trim();
      const hireDate = row.values[hireDateColIndex]?.formattedValue?.trim();

      if (!email) {
        console.log(`⏩ Skipping row ${i} - no email`);
        continue;
      }

      // Parse dates (format M/D/YYYY)
      const paidDate = lastPaidDate ? moment(lastPaidDate, 'M/D/YYYY') : null;
      const hireDateObj = hireDate ? moment(hireDate, 'M/D/YYYY') : null;

      const paidDateValid = paidDate && paidDate.isValid();
      const hireDateValid = hireDateObj && hireDateObj.isValid();

      // Qualify worker if paidDate is not valid or before 60 days ago, and hireDate is valid and before 60 days ago.
      const paidDateCondition = !paidDateValid || paidDate.isBefore(sixtyDaysAgo);
      const hireDateCondition = hireDateValid && hireDateObj.isBefore(sixtyDaysAgo);
      
      if (paidDateCondition && hireDateCondition) {
        console.log(`⚠️ Worker qualifies (access denied): ${email}`);
        workersToProcess.push({ email, rowIndex: i });
        
        // Highlight columns A–M in red
        updateRequests.push({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: i,
              endRowIndex: i + 1,
              startColumnIndex: 0,
              endColumnIndex: 13
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 1, green: 0, blue: 0 }
              }
            },
            fields: 'userEnteredFormat.backgroundColor'
          }
        });
        
        // Add checkbox in Column M (if not already there)
        updateRequests.push({
          updateCells: {
            range: {
              sheetId,
              startRowIndex: i,
              endRowIndex: i + 1,
              startColumnIndex: columnMColIndex,
              endColumnIndex: columnMColIndex + 1
            },
            rows: [{
              values: [{
                userEnteredValue: { boolValue: true }
              }]
            }],
            fields: 'userEnteredValue'
          }
        });
      } else {
        console.log(`✅ Worker does not qualify: ${email}`);
      }
    }

    // Step 2: Remove drive access for qualified workers and record rows to move
    for (const worker of workersToProcess) {
      console.log(`🔒 Removing access for: ${worker.email}`);
      await removeDriveAccess(drive, worker.email, FILE_IDS);
      confirmedRowsToMove.push(worker.rowIndex);
    }
    
    // Step 3: Move each confirmed (access-denied) row individually
    confirmedRowsToMove.sort((a, b) => b - a);
    let destinationRow = START_ROW;
    for (const rowIndex of confirmedRowsToMove) {
      // Move the row (columns A–M) to the destination row.
      updateRequests.push({
        cutPaste: {
          source: {
            sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: 0,
            endColumnIndex: 13
          },
          destination: {
            sheetId,
            rowIndex: destinationRow,
            columnIndex: 0
          }
        }
      });
      destinationRow++;
    }

    // Apply all updates
    if (updateRequests.length > 0) {
      console.log(`🔄 Applying ${updateRequests.length} update requests...`);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { requests: updateRequests }
      });
    }

    console.log('✅ Process completed successfully');
  } catch (error) {
    console.error('❌ Error:', error.message);
    if (error.response) {
      console.error('Details:', JSON.stringify(error.response.data, null, 2));
    }
    process.exit(1);
  }
}

main().catch(error => {
  console.error('❌ Unhandled error:', error);
  process.exit(1);
});

