# olook

Local Outlook CLI via COM automation. No cloud APIs, no OAuth, no telemetry.

Talks directly to Microsoft Outlook Desktop on Windows through COM — your email never leaves your machine.

## Requirements

- Windows 10/11
- Microsoft Outlook Desktop (classic `OUTLOOK.EXE`, not the new Outlook)
- Python 3.8+
- [pywin32](https://pypi.org/project/pywin32/)
- [click](https://pypi.org/project/click/)

## Install

```bash
pip install pywin32 click
git clone https://github.com/YOUR_USERNAME/olook.git
cd olook
python olook.py --help
```

## Commands

### Reading

```bash
olook inbox                        # List 10 recent emails
olook inbox -n 20 -f "Sent"       # 20 recent sent items
olook read <entry-id>              # Full email by ID
olook search "quarterly report"    # Search all fields
olook search "invoice" --field subject -n 5
```

### Composing

```bash
olook send --to "alice@example.com" --subject "Hello" --body "Hi there"
olook reply <entry-id> --body "Thanks!" --all
olook forward <entry-id> --to "bob@example.com"
```

### Organizing

```bash
olook move <entry-id> --to "Archive"
olook flag <entry-id> --text "Review by Friday"
olook mark-read <entry-id>
olook mark-read <entry-id> --unread
olook categorize <entry-id> --categories "Work,Urgent"
```

### Folders & Stats

```bash
olook folders                      # Full folder tree
olook unread                       # Unread count
olook unread -f "Sent"
olook stats                        # Top senders, date range, counts
olook scrape -n 50 --fields "subject,from,date,id"
```

### Calendar

```bash
olook cal                          # Next 7 days
olook cal -d 30                    # Next 30 days
olook cal-add --subject "Standup" --start "2025-01-15 09:00" --end "2025-01-15 09:30"
```

### Ghost Mode

Outlook must be running for COM access. Ghost mode starts it hidden at login:

```bash
olook ghost install                # Register startup task
olook ghost status                 # Check if active
olook ghost remove                 # Remove startup task
```

Without ghost mode, olook will auto-launch Outlook hidden if it's not running.

## JSON Output

Add `--json` before any command for machine-readable output:

```bash
olook --json inbox -n 5
olook --json search "project" | jq '.[].Subject'
olook --json stats
```

## How It Works

olook uses `win32com.client` to connect to Outlook's COM interface — the same API that VBA macros and Office add-ins use. No network requests, no API keys, no tokens.

On first run, it tries `GetActiveObject` to connect to an already-running Outlook (instant, zero UI). If Outlook isn't running, it launches it in `/embedding` mode (COM-server mode, no main window).

## Object Model Guard

Outlook may show "A program is trying to access email" dialogs. To suppress them, run in PowerShell (admin not required):

```powershell
$path = "HKCU:\Software\Policies\Microsoft\office\16.0\outlook\security"
New-Item -Path $path -Force | Out-Null
Set-ItemProperty -Path $path -Name "ObjectModelGuard" -Value 2 -Type DWord
Set-ItemProperty -Path $path -Name "PromptOOMAddressBookAccess" -Value 2 -Type DWord
Set-ItemProperty -Path $path -Name "PromptOOMAddressInformationAccess" -Value 2 -Type DWord
Set-ItemProperty -Path $path -Name "PromptOOMSend" -Value 2 -Type DWord
Set-ItemProperty -Path $path -Name "PromptOOMMeetingTaskRequestResponse" -Value 2 -Type DWord
Set-ItemProperty -Path $path -Name "PromptOOMSaveAs" -Value 2 -Type DWord
Set-ItemProperty -Path $path -Name "PromptOOMFormulaAccess" -Value 2 -Type DWord
```

This sets "Never warn me about suspicious activity" for COM access. Safe on single-user machines where you control what runs.

## License

MIT
