# DocuShuttle

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

A powerful email forwarding automation tool for Microsoft Outlook that helps you automatically forward emails from your Sent Items folder based on configurable filters. Useful for forwarding sent email to documentation systems where auto-forwarding is not available to the user.

## Features

- **Smart Email Filtering**: Filter emails by subject keywords, date ranges, and file number prefixes
- **Automated Forwarding**: Automatically forward matching emails to designated recipients
- **Multi-PDF Forwarding**: When an email has multiple PDF attachments with matching file number prefixes, each PDF is forwarded individually with its own file number as the subject
- **Duplicate Prevention**: Track forwarded emails to avoid sending duplicates
- **File Number Extraction**: Extract file numbers from attachments or email subjects, including optional alpha suffixes (e.g., 7607182A)
- **Subject Keyword History**: Save and recall frequently used subject keywords with typeahead autocomplete
- **Forward History Search**: Search the database for previously forwarded emails by file number or recipient
- **Preview Mode**: Preview matching emails before forwarding
- **Multi-threaded Operation**: Responsive GUI with background processing
- **Configuration Management**: Save and manage multiple recipient configurations with typeahead on recipient and keyword fields
- **Comprehensive Logging**: Detailed logging with timestamps for audit trails
- **Rate Limiting**: Configurable delays between forwarded emails
- **Auto-Update**: Automatic update checking and installation from GitHub Releases
- **Animated Splash Screen**: Professional startup experience with vortex animation

## Requirements

- Windows 10/11 (x64)
- Microsoft Outlook desktop installed and configured
- No Python installation required when using the installer

## Installation

### Using the Installer (Recommended)

1. Download the latest `DocuShuttle_Setup_v*.exe` from [GitHub Releases](https://github.com/ProcessLogicLabs/DocuShuttle/releases)
2. Run the installer - no admin privileges required
3. Launch DocuShuttle from the Start Menu or desktop shortcut

### From Source (Development)

1. **Clone the repository:**
   ```bash
   git clone https://github.com/ProcessLogicLabs/DocuShuttle.git
   cd DocuShuttle
   ```

2. **Install required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**
   ```bash
   python docushuttle.py
   ```

## Usage

### Configuration Steps

1. **Enter Forward To Email Address**:
   - Type or select the recipient email address from the dropdown
   - Typeahead autocomplete suggests previously saved recipients
   - Right-click a saved recipient to delete its configuration

2. **Set Subject Keyword**:
   - Type or select a keyword from the dropdown (e.g., "BILLING INVOICE")
   - Previously saved keywords appear in the dropdown with typeahead autocomplete
   - Right-click a saved keyword to delete it

3. **Select Date Range**:
   - Choose start and end dates for the email search

4. **Optional File Number Prefixes**:
   - Enter comma-separated numeric prefixes (e.g., "760,123") in Configuration
   - Only emails with matching file numbers will be processed
   - File numbers with trailing alpha characters (e.g., 7607182A) are supported

5. **Configure Options** (via hamburger menu > Configuration):
   - **Require Attachments**: Only forward emails with attachments
   - **Skip Previously Forwarded Emails**: Avoid forwarding duplicates
   - **Delay (Sec.)**: Add delay between forwarded emails

### Operations

#### Preview Mode
Click **Preview** to see a list of emails that match your criteria without forwarding them. For emails with multiple matching PDFs, each PDF appears as a separate row.

#### Scan and Forward
Click **Scan and Forward** to automatically forward all matching emails to the configured recipient.

#### Cancel Operation
Click **Cancel** to stop an ongoing search or forward operation.

#### Forward History
Access via the hamburger menu > **Forward History...** to search the database for previously forwarded emails. Filter by file number or recipient (partial match supported).

## How It Works

### Email Scanning Process

When you click **Preview** or **Scan and Forward**, the application performs the following steps:

1. **Connect to Outlook**: Establishes a COM connection to Microsoft Outlook using the Windows API
2. **Access Sent Items**: Opens your Sent Items folder and retrieves the email count
3. **Apply Subject Filter**: Uses Outlook's MAPI filter to find emails containing your subject keyword (case-insensitive)
4. **Date Range Filtering**: Checks each email's sent date against your specified date range
5. **File Number Extraction**: If prefixes are specified, extracts file numbers from attachment filenames or email subjects (including optional trailing alpha characters like A, B, etc.)
6. **Duplicate Check**: If "Skip Previously Forwarded" is enabled, checks the database for previously forwarded file numbers
7. **Attachment Check**: If "Require Attachments" is enabled, skips emails without attachments

### Forward Process

#### Single-Attachment Emails
1. Creates a forward copy of the matching email
2. Sets the recipient to your configured "Forward To" address
3. Replaces the subject line with the extracted file number (alpha suffix stripped)
4. Sends the forwarded email
5. Logs the file number to the database to prevent future duplicates

#### Multi-PDF Emails
When an email has **two or more PDF attachments** whose filenames match a configured prefix:

1. Creates a separate forward for **each** matching PDF
2. Each forward contains **only that one PDF** attachment (all others are removed)
3. Sets the subject line to the PDF's file number (alpha suffix stripped)
4. Logs each file number individually for duplicate tracking
5. Applies the configured delay between each forward

**Example**: An email with attachments `7607182A.pdf`, `7607189A.pdf`, `7607303A.pdf` produces three separate forwards with subjects `7607182`, `7607189`, `7607303`.

### Configuration Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| **Forward To** | Email address where matching emails will be forwarded | Required |
| **Subject Keyword** | Text to search for in email subjects (case-insensitive) | "BILLING INVOICE" |
| **Start Date** | Beginning of the date range to search | Today |
| **End Date** | End of the date range to search | Today |
| **File Number Prefixes** | Comma-separated numeric prefixes (e.g., "760,123") to filter and extract file numbers | Empty (all emails) |
| **Delay (Sec.)** | Seconds to wait between forwarding each email | 0 |
| **Require Attachments** | Only forward emails that have attachments | Checked |
| **Skip Previously Forwarded** | Skip emails with file numbers already in the tracking database | Checked |

### File Number Extraction

File numbers are extracted using the following priority:

1. **From Attachments**: Scans attachment filenames for patterns matching your prefixes (e.g., `7607182A.pdf` extracts `7607182A`)
2. **From Subject**: If no attachment match, searches the email subject for matching patterns

The pattern matched is: `{prefix}` + digits to fill 7 characters total + optional trailing alpha (A-Z).

When used as the forwarded email subject, the trailing alpha character is stripped (e.g., `7607182A` becomes `7607182`). The full file number including the alpha is retained in the database for accurate duplicate tracking.

### Rate Limiting

To prevent overwhelming the mail server:

- **Manual Delay**: Configure any delay in the Configuration dialog
- **Automatic Delay**: A 3-second minimum delay is automatically applied when the date range exceeds 8 days
- **Recommended**: Use 1-3 second delays for large batches of emails

## Database

The application uses a SQLite database (`docushuttle.db`) stored in `%LOCALAPPDATA%\DocuShuttle\`.

### Clients Table
Stores configuration for each recipient:
- recipient (email address)
- start_date, end_date
- file_number_prefix
- subject_keyword
- require_attachments, skip_forwarded
- delay_seconds

### ForwardedEmails Table
Tracks forwarded emails to prevent duplicates:
- file_number (includes alpha suffix if present)
- recipient
- forwarded_at (timestamp)

### Settings Table
Application-level key/value settings:
- last_used_email, last_update_check, auto_update

### Keywords Table
Saved subject keywords for the keyword combobox dropdown.

## Logging

The application maintains logging visible in the **Log** tab with real-time updates during operations. All timestamps use US/Eastern timezone.

Error details are also written to `error.log` in `%LOCALAPPDATA%\DocuShuttle\` for diagnostics.

## Safety Features

- **Thread-Safe Operations**: All database and GUI operations are thread-safe
- **Retry Logic**: Automatic retry for Outlook connection failures
- **Input Validation**: Email format and date validation
- **Filter Sanitization**: Protection against MAPI filter injection
- **Per-Attachment Error Handling**: In multi-PDF mode, a failure on one attachment does not stop the remaining attachments
- **Cancel Support**: Operations can be cancelled mid-batch, including during multi-PDF forwarding

## Date Range Behavior

**Important**: If you select a date range exceeding 8 days, the application automatically applies a 3-second delay between forwarded emails to prevent Outlook throttling.

## Auto-Update

- DocuShuttle checks GitHub for new releases periodically
- When an update is available, you are prompted to download and install
- Updates are downloaded to `%LOCALAPPDATA%\DocuShuttle\data\updates\`
- No forced or silent updates - you always approve before installing

## Troubleshooting

### Outlook Connection Issues
- Ensure Microsoft Outlook desktop is installed and configured
- Try closing and reopening Outlook
- Check if Outlook is set as the default mail client
- Verify both DocuShuttle and Outlook are 64-bit

### No Emails Found
- Verify the subject keyword matches emails in Sent Items
- Adjust the date range to include relevant emails
- Check if "Require Attachments" should be unchecked
- Disable "Skip Previously Forwarded Emails" to resend

### Multi-PDF Not Splitting
- Ensure file number prefixes are configured in Configuration
- The email must have **2 or more** PDF attachments with filenames matching a prefix
- Single-PDF emails forward normally without splitting

### "Failed to Load Python DLL" Error
- This is resolved in v1.7.1+ which bundles all DLLs in the install directory
- Reinstall using the latest installer from GitHub Releases
- If the issue persists, check if antivirus is quarantining files in the install directory

### Slow Performance
- Reduce the date range for faster searching
- Use more specific subject keywords
- Close other applications to free up resources

## Building

### Build Executable (PyInstaller)
```bash
pip install pyinstaller
pyinstaller docushuttle.spec
```
Output: `dist/DocuShuttle/` directory with exe and all dependencies.

### Build Installer (Inno Setup 6+)
```bash
iscc DocuShuttle_Setup.iss
```
Output: `dist/DocuShuttle_Setup_v{version}.exe`

### CI/CD
Pushing a version tag (e.g., `v1.7.3`) triggers a GitHub Actions workflow that builds the installer and creates a GitHub Release automatically.

## Project Structure

```
DocuShuttle/
├── docushuttle.py            # Main application file
├── docushuttle.spec          # PyInstaller specification (onedir mode)
├── DocuShuttle_Setup.iss     # Inno Setup installer script
├── create_icon.py            # Icon generation script
├── myicon.ico                # Application icon (multi-resolution)
├── myicon.png                # PNG version of icon
├── requirements.txt          # Python dependencies
├── README.md                 # This file
├── IT_DEPLOYMENT_GUIDE.md    # IT department technical summary
├── .github/workflows/
│   └── build.yml             # GitHub Actions CI/CD workflow
└── dist/                     # Build output directory
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see below for details:

```
MIT License

Copyright (c) 2024 Royal Payne

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## Author

**Royal Payne** - [Process Logic Labs](https://github.com/ProcessLogicLabs)

## Acknowledgments

- Built with [pywin32](https://github.com/mhammond/pywin32) for Outlook integration
- GUI built with [PyQt5](https://www.riverbankcomputing.com/software/pyqt/)

## Support

For issues, questions, or suggestions, please open an issue on the [GitHub repository](https://github.com/ProcessLogicLabs/DocuShuttle/issues).

---

**Note**: This tool is designed for Windows environments with Microsoft Outlook. It accesses your Outlook Sent Items folder and requires appropriate permissions to send emails on your behalf.
