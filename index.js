require('dotenv').config();
const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const open = require('open');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
const LOG_FILE = path.join(process.cwd(), 'gtsignups.log');

// Function to write logs
async function writeLog(message) {
    const timestamp = new Date().toISOString();
    const logMessage = `${timestamp}: ${message}\n`;
    console.log(message);
    await fs.appendFile(LOG_FILE, logMessage);
}

async function loadSavedCredentialsIfExist() {
    try {
        const content = await fs.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        await writeLog('Loaded saved credentials');
        return google.auth.fromJSON(credentials);
    } catch (err) {
        await writeLog('No saved credentials found');
        return null;
    }
}

async function saveCredentials(client) {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
    });
    await fs.writeFile(TOKEN_PATH, payload);
    await writeLog('Credentials saved');
}

async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
        return client;
    }
    client = await authenticate({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
    });
    if (client.credentials) {
        await saveCredentials(client);
    }
    await writeLog('New authorization completed');
    return client;
}

async function queryEmails(auth) {
    await writeLog('Querying emails for Growth Track Sign Ups...');
    const gmail = google.gmail({ version: 'v1', auth });

    // Calculate date range (7 days ago to now)
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    const after = Math.floor(sevenDaysAgo.getTime() / 1000);

    const res = await gmail.users.messages.list({
        userId: 'me',
        q: `subject:"Growth Track Signup" after:${after}`,
    });
    const messages = res.data.messages || [];
    await writeLog(`Found ${messages.length} relevant emails from the last 7 days`);
    return messages;
}

async function extractSignupInfo(auth, messageId) {
    await writeLog(`Extracting signup information from email ${messageId}...`);
    const gmail = google.gmail({ version: 'v1', auth });
    const res = await gmail.users.messages.get({
        userId: 'me',
        id: messageId,
        format: 'full',
    });

    // Get email body
    const body = res.data.payload.parts ?
        Buffer.from(res.data.payload.parts[0].body.data, 'base64').toString('utf8') :
        Buffer.from(res.data.payload.body.data, 'base64').toString('utf8');

    await writeLog('Parsing email body...');
    await writeLog(`Raw email body: ${body}`);

    // Updated regex patterns to match HTML table structure
    const patterns = {
        name: /<td valign="top">Full Name:<\/td>\s*<td>([^<]+)<\/td>/i,
        phone: /<td valign="top">Phone:<\/td>\s*<td>([^<]+)<\/td>/i,
        email: /<td valign="top">Email:<\/td>\s*<td>([^<]+)<\/td>/i
    };

    const extractedInfo = {};

    // Extract information using updated patterns
    for (const [key, pattern] of Object.entries(patterns)) {
        const match = body.match(pattern);
        if (match) {
            extractedInfo[key] = match[1].trim();
            await writeLog(`Extracted ${key}: ${extractedInfo[key]}`);
        } else {
            extractedInfo[key] = '';
            await writeLog(`Warning: Could not extract ${key} from email body`);
        }
    }

    // Use email received date as registration date
    const emailDate = new Date(parseInt(res.data.internalDate));
    extractedInfo.date = formatDate(emailDate);
    await writeLog(`Using email date as registration date: ${extractedInfo.date}`);

    // Validate extracted information
    if (!extractedInfo.name || !extractedInfo.email || !extractedInfo.phone) {
        await writeLog('Warning: Some required fields are missing');
        await writeLog(`Extracted information: ${JSON.stringify(extractedInfo)}`);
    }

    return {
        date: extractedInfo.date,
        name: extractedInfo.name || 'Not provided',
        phone: extractedInfo.phone || 'Not provided',
        email: extractedInfo.email || 'Not provided'
    };
}

// Update the formatDate function to be more specific
function formatDate(date) {
    const options = {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        timeZone: 'Africa/Johannesburg' // Set to South African timezone
    };
    return date.toLocaleDateString('en-ZA', options);
}

async function createOrGetSpreadsheet(auth) {
    const drive = google.drive({ version: 'v3', auth });
    const sheets = google.sheets({ version: 'v4', auth });

    // First, check if "Growth Track Registrations" folder exists
    const folderResponse = await drive.files.list({
        q: "mimeType='application/vnd.google-apps.folder' and name='Growth Track Registrations'",
        fields: 'files(id, name)',
    });

    let folderId;
    if (folderResponse.data.files.length > 0) {
        folderId = folderResponse.data.files[0].id;
        await writeLog(`Found existing "Growth Track Registrations" folder with ID: ${folderId}`);
    } else {
        const folderMetadata = {
            name: 'Growth Track Registrations',
            mimeType: 'application/vnd.google-apps.folder',
        };
        const folder = await drive.files.create({
            resource: folderMetadata,
            fields: 'id',
        });
        folderId = folder.data.id;
        await writeLog(`Created new "Growth Track Registrations" folder with ID: ${folderId}`);
    }

    // Check for existing spreadsheet named "Growth Track Signups"
    const spreadsheetResponse = await drive.files.list({
        q: `'${folderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name='Growth Track Signups'`,
        fields: 'files(id, name)',
    });

    let spreadsheetId;
    let sheetId;

    if (spreadsheetResponse.data.files.length > 0) {
        // Use existing spreadsheet
        spreadsheetId = spreadsheetResponse.data.files[0].id;
        await writeLog(`Found existing spreadsheet with ID: ${spreadsheetId}`);

        // Get the sheet ID
        const spreadsheetInfo = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId
        });
        sheetId = spreadsheetInfo.data.sheets[0].properties.sheetId;
    } else {
        // Create new spreadsheet
        const spreadsheet = await sheets.spreadsheets.create({
            resource: {
                properties: {
                    title: 'Growth Track Signups',
                },
                sheets: [{
                    properties: {
                        title: 'Sheet1',
                        gridProperties: {
                            rowCount: 1000,
                            columnCount: 4
                        }
                    },
                    data: [{
                        rowData: [{
                            values: [
                                { userEnteredValue: { stringValue: 'Registration Date' } },
                                { userEnteredValue: { stringValue: 'Full Name' } },
                                { userEnteredValue: { stringValue: 'Phone' } },
                                { userEnteredValue: { stringValue: 'Email' } }
                            ]
                        }]
                    }]
                }]
            },
        });

        spreadsheetId = spreadsheet.data.spreadsheetId;
        sheetId = spreadsheet.data.sheets[0].properties.sheetId;
        await writeLog(`Created new spreadsheet with ID: ${spreadsheetId}`);

        // Format header row
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: spreadsheetId,
            resource: {
                requests: [
                    {
                        repeatCell: {
                            range: {
                                sheetId: sheetId,
                                startRowIndex: 0,
                                endRowIndex: 1,
                                startColumnIndex: 0,
                                endColumnIndex: 4
                            },
                            cell: {
                                userEnteredFormat: {
                                    backgroundColor: { red: 0.9, green: 0.9, blue: 0.9 },
                                    textFormat: { bold: true },
                                    horizontalAlignment: 'CENTER'
                                }
                            },
                            fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                        }
                    },
                    {
                        updateDimensionProperties: {
                            range: {
                                sheetId: sheetId,
                                dimension: 'COLUMNS',
                                startIndex: 0,
                                endIndex: 4
                            },
                            properties: {
                                pixelSize: 200
                            },
                            fields: 'pixelSize'
                        }
                    }
                ]
            }
        });

        // Move the spreadsheet to the folder
        await drive.files.update({
            fileId: spreadsheetId,
            addParents: folderId,
            fields: 'id, parents',
        });
        await writeLog(`Moved spreadsheet to "Growth Track Registrations" folder`);
    }

    return spreadsheetId;
}

async function saveToGoogleSheets(auth, data) {
    await writeLog('Saving information to Google Sheets...');
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = await createOrGetSpreadsheet(auth);
    const range = 'Sheet1!A:D'; // Updated range to match four columns

    const response = await sheets.spreadsheets.values.append({
        spreadsheetId,
        range,
        valueInputOption: 'USER_ENTERED',
        resource: {
            values: [[data.date, data.name, data.phone, data.email]], // Include date in the first column
        },
    });
    await writeLog('Information saved to Google Sheets');
    return { spreadsheetId, response: response.data };
}

async function sendEmail(auth, spreadsheetId) {
    await writeLog('Preparing to send email with attachment...');
    const sheets = google.sheets({ version: 'v4', auth });
    const drive = google.drive({ version: 'v3', auth });

    await writeLog('Exporting spreadsheet as Excel...');
    const res = await drive.files.export({
        fileId: spreadsheetId,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }, { responseType: 'arraybuffer' });

    await writeLog('Saving Excel file temporarily...');
    await fs.writeFile('GrowthTrackSignups.xlsx', Buffer.from(res.data));

    await writeLog('Creating email transporter...');
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS,
        },
    });

    await writeLog('Sending email...');
    const info = await transporter.sendMail({
        from: process.env.EMAIL_USER,
        to: process.env.RECIPIENT_EMAIL,
        subject: 'Growth Track Signups',
        text: 'Please find attached the latest Growth Track signups.',
        attachments: [
            {
                filename: 'GrowthTrackSignups.xlsx',
                path: './GrowthTrackSignups.xlsx',
            },
        ],
    });

    await writeLog(`Email sent: ${info.messageId}`);

    await writeLog('Deleting temporary Excel file...');
    await fs.unlink('GrowthTrackSignups.xlsx');
}

async function main() {
    try {
        await writeLog('Starting the Growth Track Signup process...');
        const auth = await authorize();
        const messages = await queryEmails(auth);

        let spreadsheetId;
        for (const message of messages) {
            await writeLog(`Processing message ${message.id}...`);
            const signupInfo = await extractSignupInfo(auth, message.id);
            const result = await saveToGoogleSheets(auth, signupInfo);
            spreadsheetId = result.spreadsheetId;
        }

        if (spreadsheetId) {
            await sendEmail(auth, spreadsheetId);
        } else {
            await writeLog('No new signups to process');
        }

        await writeLog('Process completed successfully');
    } catch (error) {
        await writeLog(`An error occurred: ${error.message}`);
        await writeLog(`Error details: ${JSON.stringify(error)}`);
    }
}

main();
