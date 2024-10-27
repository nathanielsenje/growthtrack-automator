 # Growth Track Signup Processor

This application automates the processing of Growth Track signup emails by extracting registration information, storing it in Google Sheets, and sending summary reports via email.

## Features

- Automatically monitors Gmail inbox for new Growth Track signup emails
- Extracts registration details (name, phone, email) from email content
- Stores signup information in a Google Sheets spreadsheet
- Marks processed emails as read
- Sends summary reports with Excel attachments
- Prevents duplicate entries via email tracking
- Comprehensive logging system

## Prerequisites

Before running this application, you need:

1. Node.js installed on your system
2. A Google Cloud Project with the following APIs enabled:
   - Gmail API
   - Google Sheets API
   - Google Drive API
3. OAuth 2.0 credentials (credentials.json)
4. Gmail account with appropriate permissions

## Environment Variables

Create a `.env` file in the root directory with the following variables:
EMAIL_USER=your.email@gmail.com
EMAIL_PASS=your-app-specific-password
RECIPIENT_EMAIL=recipient@example.com
SPREADSHEET_ID=your-google-spreadsheet-id

## Installation

1. Clone the repository:
git clone https://github.com/yourusername/growth-track-processor.git
cd growth-track-processor

2. Install dependencies:
npm install

3. Set up Google Cloud credentials:
   - Create a project in Google Cloud Console
   - Enable required APIs
   - Create OAuth 2.0 credentials
   - Download credentials as `credentials.json`
   - Place `credentials.json` in the project root

4. Configure environment variables:
   - Copy `.env.example` to `.env`
   - Fill in your configuration details

## Usage

Run the application:
npm start


The application will:
1. Check for new Growth Track signup emails
2. Process unread emails and extract registration information
3. Save the data to Google Sheets
4. Generate an Excel report
5. Send the report via email
6. Mark processed emails as read
7. Log all processed emails to prevent duplicates

## File Structure

- `index.js`: Main application script
- `gtsignups.log`: Log file for application logs
- `token.json`: OAuth token file
- `emails.json`: Email tracking file
- `package.json`: Node.js dependencies and scripts
- `README.md`: This documentation file

## Logging

The application maintains detailed logs in the `logs` directory. Each log entry includes:
- Timestamp
- Operation type
- Success/failure status
- Error details (if any)

## Error Handling

The application includes comprehensive error handling for:
- Email processing failures
- Google API connectivity issues
- File system operations
- Email sending errors

All errors are logged and will not crash the application.

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support, please contact the development team or create an issue in the repository.