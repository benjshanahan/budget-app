const { google } = require('googleapis');

const SPREADSHEET_ID = '114vo9a8wrVAsDTU6IJASsADSI6Gtv8Dlj2vSuaDwwPg';

const NDIS_FOLDERS = {
  core: '1upqz4gwIM0q7klvfkTIAhNkS3ecgs4US',
  idls: '1GVmWlO4dSzhvPZnzG4hhiHaZA9-d2yKj',
  sc:   '1cKzi9eZ6hLtmbAsgYQKWTqTmCII2pPa0',
  at:   '1T4TRV-23F87jVO6b2gIqKyiWiypxTmTM',
};

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
    ],
  });
  return auth;
}

exports.handler = async (event) => {
  // Handle CORS preflight
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      },
      body: '',
    };
  }

  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json',
  };

  try {
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    const body = event.body ? JSON.parse(event.body) : {};
    const action = body.action || event.queryStringParameters?.action;

    if (action === 'read') {
      const ranges = body.ranges || [
        'Transactions!A:K',
        'Config!A:B',
        'MoneyOwed!A:C',
        'Spending!A:B',
        'Workflow!A:B',
        'DeletedTx!A:K',
        'Cycles!A:F',
        'NDIS!A:I',
        'CSVPending!A:A',
      ];

      const response = await sheets.spreadsheets.values.batchGet({
        spreadsheetId: SPREADSHEET_ID,
        ranges,
      });

      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ success: true, data: response.data.valueRanges }),
      };
    }

    if (action === 'write') {
      const data = body.data;
      // Clear each range before writing to remove stale rows
      await sheets.spreadsheets.values.batchClear({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          ranges: data.map(d => d.range),
        },
      });
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: 'RAW',
          data,
        },
      });

      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ success: true }),
      };
    }

    if (action === 'listDriveFiles') {
      const drive = google.drive({ version: 'v3', auth });
      const results = {};
      for (const [catId, folderId] of Object.entries(NDIS_FOLDERS)) {
        const resp = await drive.files.list({
          q: `'${folderId}' in parents and trashed = false`,
          fields: 'files(id, name, webViewLink)',
          pageSize: 200,
        });
        results[catId] = resp.data.files;
      }
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ success: true, files: results }),
      };
    }

    if (action === 'uploadToDrive') {
      const { catId, filename, fileBase64, mimeType } = body;
      const folderId = NDIS_FOLDERS[catId];
      if (!folderId) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Unknown category' }) };

      const drive = google.drive({ version: 'v3', auth });
      const buffer = Buffer.from(fileBase64, 'base64');

      const response = await drive.files.create({
        requestBody: {
          name: filename,
          parents: [folderId],
        },
        media: {
          mimeType: mimeType || 'application/pdf',
          body: require('stream').Readable.from(buffer),
        },
        fields: 'id, name, webViewLink',
      });

      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          success: true,
          fileId: response.data.id,
          fileName: response.data.name,
          webViewLink: response.data.webViewLink,
        }),
      };
    }

    if (action === 'ensureSheets') {
      const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
      const existing = meta.data.sheets.map(s => s.properties.title);
      const needed = ['Transactions', 'Config', 'MoneyOwed', 'Spending', 'Workflow', 'DeletedTx', 'Cycles', 'NDIS', 'CSVPending'];
      const toCreate = needed.filter(n => !existing.includes(n));

      if (toCreate.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: {
            requests: toCreate.map(title => ({ addSheet: { properties: { title } } })),
          },
        });
      }

      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ success: true, created: toCreate }),
      };
    }

    return {
      statusCode: 400,
      headers,
      body: JSON.stringify({ error: 'Unknown action' }),
    };

  } catch (err) {
    console.error('Sheets error:', err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
