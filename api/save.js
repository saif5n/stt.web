const { google } = require('googleapis');

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_APPLICATION_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ message: 'Method not allowed' });
  
  const { rowId, username, url, platform, judgement, notes } = req.body;
  const spreadsheetId = '170I3V2-Hl1KwTrnvIKtlpkxZeMvg8xXYKK2HjJtnnY0'; 
  
  const sheets = google.sheets({ version: 'v4', auth });
  const isSkipped = (judgement === 'Skipped');

  try {
    // 1. UPDATE URLs sheet
    if (isSkipped) {
      // If skipped, update Status (Col D) and Reason (Col E)
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `URLs!D${rowId}:E${rowId}`, 
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [['Skipped', notes]] }, 
      });
    } else {
      // If reviewed, just update Status (Col D)
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `URLs!D${rowId}`, 
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [['Reviewed']] }, 
      });
    }

    // 2. INSERT to Submissions sheet (Only if NOT skipped)
    if (!isSkipped) {
      await sheets.spreadsheets.values.insert({
        spreadsheetId,
        range: 'Submissions!A2', 
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { 
          values: [[
            new Date().toLocaleString(), 
            username, 
            url, 
            platform, 
            judgement, 
            notes
          ]] 
        },
      });
    }

    return res.status(200).json({ success: true });
  } catch (error) {
    console.error("Critical Save Error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
}