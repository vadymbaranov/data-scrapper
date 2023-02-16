import express from 'express';
import { google }from 'googleapis';
import DOMParser from 'dom-parser';

const app = express();
app.use(express.json());

app.listen(5000, () => console.log('Server is running on PORT 5000'));

const authentication = async () => {
  const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json',
    scopes: 'https://www.googleapis.com/auth/spreadsheets'
  });

  const client = await auth.getClient();

  const sheets = google.sheets({
    version: 'v4',
    auth: client
  });

  return { sheets };
}

const spreadsheetId = '1R-PkXbc-mhkURUl8A0BKwT9c5dqaQ5RHaFBSBlRXs7w';

app.get('/', async (req, res) => {
  try {
    const { sheets } = await authentication();

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: 'Sheet1',
    })
    res.send(response.data)
  } catch (err) {
    console.log(err);
    res.sendStatus(500);
  }
})

app.post('/', async (req, res) => {
  try {
    const { newName, newValue } = req.body;
    const { sheets } = await authentication();
    const writeReq = await sheets.spreadsheets.values.append({
      spreadsheetId: spreadsheetId,
      range: 'Sheet1',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [
          [newName, newValue],
        ]
      }
    });

    if (writeReq.status === 200) {
      return res.json({ msg: 'Spreadsheet updated successfully!' });
    }

    return res.json({ msg: 'Something went wrong while updating the spreadsheet.' });
  } catch(err) {
    console.log('Error updating the spreadsheet', err);
    res.sendStatus(500);
  }
})

async function importGridData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const url = 'https://www.crunchbase.com/discover/organization.companies';
    const response = UrlFetchApp.fetch(url);
    const content = response.getContentText();
    const parser = new DOMParser();
    const doc = parser.parseFromString(content, 'text/html');
    const grid = doc.querySelector('sheet-grid');
    const rows = grid.querySelectorAll('grid-row');
    const data = [];
    for (let i = 0; i < rows.length; i++) {
      const row = [];
      const cells = rows[i].querySelectorAll('grid-cell');
      for (let j = 0; j < cells.length; j++) {
        row.push(cells[j].textContent.trim());
      }
      data.push(row);
    }
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  } catch(err) {
    return 'Unable to parse data, check your credentials';
  }
}

importGridData()
  .then((sheet) => console.log(sheet));

