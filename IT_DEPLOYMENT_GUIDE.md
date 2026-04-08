# DocuShuttle v1.7.3 - IT Technical Summary

## What It Does
DocuShuttle automates forwarding emails from a user's Microsoft Outlook **Sent Items** folder to designated recipients. Users configure filters (subject keyword, date range, file number prefixes) and the application finds matching emails and forwards them automatically. This eliminates manual forwarding of billing invoices and customs documents.

## Deployment Model
- **Installer:** `DocuShuttle_Setup_v1.7.3.exe` (Inno Setup, ~45 MB)
- **Install location:** `%PROGRAMFILES%\DocuShuttle\` (or user-chosen)
- **Privileges:** Runs at **lowest** privilege - no admin required to install or run
- **Architecture:** x64 Windows only

## What Gets Installed
| Path | Contents |
|------|----------|
| `{install dir}\DocuShuttle.exe` | Main application |
| `{install dir}\*.dll` | Python 3.11 runtime, Qt5, pywin32 DLLs (self-contained, no Python install needed) |
| `{install dir}\myicon.ico/png` | Application icon |
| `%LOCALAPPDATA%\DocuShuttle\` | User data directory (created at runtime) |

## Data Storage
| File | Location | Purpose |
|------|----------|---------|
| `docushuttle.db` | `%LOCALAPPDATA%\DocuShuttle\` | SQLite database - saved configs, forwarded email tracking |
| `settings.json` | `%LOCALAPPDATA%\DocuShuttle\` | Update check timestamps |
| `error.log` | `%LOCALAPPDATA%\DocuShuttle\` | Error/diagnostic log |

## Network & COM Access
| Resource | Protocol | Purpose |
|----------|----------|---------|
| Microsoft Outlook | COM/MAPI (`win32com`) | Read Sent Items, create and send forwards |
| `api.github.com` | HTTPS (outbound, port 443) | Check for application updates (once per hour) |
| `github.com` | HTTPS (outbound, port 443) | Download update installers |

**No other network connections are made.** No data is sent externally - all email operations go through the user's existing Outlook/Exchange connection.

## Outlook Integration Details
- Uses **Windows COM automation** (`win32com.client`) to interface with Outlook
- Accesses only the **Sent Items** folder (MAPI folder index 5)
- Operations: read email metadata, create forward copies, send via user's account
- Requires Outlook desktop app to be installed and configured
- Runs as the logged-in user's Outlook identity - no separate credentials

## Security Considerations
- **No elevated privileges** - installs and runs as standard user
- **No credential storage** - uses existing Outlook session
- **No external data transmission** - only GitHub API for update checks
- **SQLite database** is local only, contains file numbers and recipient emails
- **DLLs are bundled** in the install directory (not extracted to temp) - avoids antivirus false positives
- **No registry modifications** beyond standard Inno Setup uninstall entries

## System Requirements
- Windows 10/11 (x64)
- Microsoft Outlook desktop (configured with an Exchange/M365 account)
- ~100 MB disk space
- No Python installation required (runtime is bundled)

## Auto-Update Mechanism
- Checks `api.github.com/repos/ProcessLogicLabs/DocuShuttle/releases/latest` periodically
- Compares remote version to installed version
- Downloads installer `.exe` to `%LOCALAPPDATA%\DocuShuttle\data\updates\`
- Prompts user before installing - no silent/forced updates

## Firewall Rules (if needed)
- **Allow outbound HTTPS** to `api.github.com` and `github.com` for auto-updates
- If updates are managed internally, these can be blocked - the app functions fully without them
