require('dotenv').config();
const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const open = require('open');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive'];
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

// Update the queryEmails function to use the correct subject and add better logging
async function queryEmails(auth) {
    try {
        const gmail = google.gmail({ version: 'v1', auth });

        // Calculate date range (7 days ago to now)
        const sevenDaysAgo = new Date();
        sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
        const after = Math.floor(sevenDaysAgo.getTime() / 1000);

        await writeLog('Querying Gmail for unread Growth Track Sign Up emails...');

        // First try with exact subject
        let response = await gmail.users.messages.list({
            userId: 'me',
            q: `subject:"Growth Track Signup" is:unread`,
            maxResults: 100
        });

        let messages = response.data.messages || [];
        await writeLog(`Found ${messages.length} unread messages with subject "Growth Track Sign Up"`);

        // If no messages found, try alternative subject
        if (messages.length === 0) {
            response = await gmail.users.messages.list({
                userId: 'me',
                q: `subject:"Growth Track Sign Up Form" is:unread`,
                maxResults: 100
            });
            messages = response.data.messages || [];
            await writeLog(`Found ${messages.length} unread messages with subject "Growth Track Sign Up Form"`);
        }

        // Log details about found messages
        if (messages.length > 0) {
            await writeLog('Getting message details...');
            for (const message of messages) {
                const details = await gmail.users.messages.get({
                    userId: 'me',
                    id: message.id,
                    format: 'metadata',
                    metadataHeaders: ['Subject']
                });
                const subject = details.data.payload.headers.find(h => h.name === 'Subject')?.value;
                await writeLog(`Message ${message.id} has subject: "${subject}"`);
            }
        }

        return messages;
    } catch (error) {
        await writeLog(`Error querying emails: ${error.message}`);
        throw error;
    }
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
    const tempFilePath = path.join(process.cwd(), 'GrowthTrackSignups.xlsx');
    try {
        await writeLog('Preparing to send email with attachment...');
        const drive = google.drive({ version: 'v3', auth });

        await writeLog('Exporting spreadsheet as Excel...');
        const res = await drive.files.export({
            fileId: spreadsheetId,
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }, { responseType: 'arraybuffer' });

        await writeLog('Saving Excel file temporarily...');
        await fs.writeFile(tempFilePath, Buffer.from(res.data));

        const sevenDaysAgo = new Date();
        sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
        const dateStr = formatDate(sevenDaysAgo);

        const emailBody = `Hello,

I trust you are well. Please see the attached file for the latest growth track registrations, from ${dateStr}

Best Regards,
Nathaniel Senje
Digital Content Manager and Music Director
+255 747 428 797
The Ocean International Community Church (TAG),
Plot No. 1831, The Little Theatre`;

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
            text: emailBody,
            attachments: [
                {
                    filename: 'GrowthTrackSignups.xlsx',
                    path: tempFilePath,
                }
            ],
        });

        await writeLog(`Email sent: ${info.messageId}`);

        // Clean up temp file
        await writeLog('Cleaning up temporary file...');
        await fs.unlink(tempFilePath);
    } catch (error) {
        // Clean up temp file even if there's an error
        try {
            await fs.unlink(tempFilePath);
        } catch {
            // Ignore cleanup errors in error handler
        }
        throw error;
    }
}

// Update the markEmailAsRead function with better error handling
async function markEmailAsRead(auth, messageId) {
    try {
        const gmail = google.gmail({ version: 'v1', auth });
        await writeLog(`Marking email ${messageId} as read...`);
        
        await gmail.users.messages.modify({
            userId: 'me',
            id: messageId,
            resource: {
                removeLabelIds: ['UNREAD']
            }
        });
        
        await writeLog(`Successfully marked email ${messageId} as read`);
    } catch (error) {
        await writeLog(`Error marking email as read: ${error.message}`);
        throw error;
    }
}

// Add this function to check for duplicate entries
async function isDuplicateEntry(auth, data, spreadsheetId) {
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: 'Sheet1!A:D',
    });

    const rows = response.data.values || [];
    return rows.some(row =>
        row[1] === data.name &&
        row[2] === data.phone &&
        row[3] === data.email
    );
}

// Modify the main function to use these new features
async function main() {
    try {
        await writeLog('Starting the Growth Track Signup process...');
        const auth = await authorize();
        const messages = await queryEmails(auth);

        let spreadsheetId;
        let processedCount = 0;
        let skippedCount = 0;

        await writeLog(`Processing ${messages.length} messages...`);

        for (const message of messages) {
            try {
                await writeLog(`Processing message ${message.id}...`);
                const signupInfo = await extractSignupInfo(auth, message.id);

                // Get or create spreadsheet if this is the first entry
                if (!spreadsheetId) {
                    spreadsheetId = await createOrGetSpreadsheet(auth);
                }

                // Check for duplicates before saving
                const isDuplicate = await isDuplicateEntry(auth, signupInfo, spreadsheetId);
                if (isDuplicate) {
                    await writeLog(`Skipping duplicate entry for ${signupInfo.name}`);
                    skippedCount++;
                } else {
                    const result = await saveToGoogleSheets(auth, signupInfo);
                    spreadsheetId = result.spreadsheetId;
                    processedCount++;
                    await writeLog(`Successfully processed signup for ${signupInfo.name}`);
                }

                // Mark as read regardless of whether it was a duplicate or not
                await markEmailAsRead(auth, message.id);
            } catch (error) {
                await writeLog(`Error processing message ${message.id}: ${error.message}`);
                // Continue with next message even if one fails
                continue;
            }
        }

        await writeLog(`Processing complete. ${processedCount} new entries added, ${skippedCount} duplicates skipped`);

        if (processedCount > 0) {
            await writeLog(`Sending email with ${processedCount} new signups`);
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
