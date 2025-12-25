# gog

Google in your terminal - CLI for Gmail, Calendar, Drive, Contacts, Tasks, and Sheets.

## Features

- **Gmail** - search threads, send emails, manage labels, drafts, filters, delegation, vacation settings, and watch (Pub/Sub push)
- **Calendar** - list/create/update events, detect conflicts, manage invitations, check free/busy status
- **Drive** - list/search/upload/download files, manage permissions, organize folders
- **Contacts** - search/create/update contacts, access Workspace directory
- **Tasks** - manage tasklists and tasks: create/add/update/done/undo/delete/clear
- **Sheets** - read/write/update spreadsheets, create new sheets
- **People** - access profile information
- **Multiple account support** - manage multiple Google accounts simultaneously
- **Secure credential storage** using OS keyring (Keychain on macOS, Secret Service on Linux, Credential Manager on Windows)
- **Auto-refreshing tokens** - authenticate once, use indefinitely
- **Parseable output** - JSON mode for scripting and automation

## Installation

### Homebrew

```bash
brew install steipete/tap/gogcli
```

### Build from Source

```bash
git clone https://github.com/steipete/gogcli.git
cd gogcli
make
```

Run:

```bash
./bin/gog --help
```

## Quick Start

### 1. Get OAuth2 Credentials

Before adding an account, create OAuth2 credentials from Google Cloud Console:

1. Create a project: https://console.cloud.google.com/projectcreate
2. Enable the APIs you need:
   - Gmail API: https://console.cloud.google.com/apis/api/gmail.googleapis.com
   - Google Calendar API: https://console.cloud.google.com/apis/api/calendar-json.googleapis.com
   - Google Drive API: https://console.cloud.google.com/apis/api/drive.googleapis.com
   - People API (Contacts): https://console.cloud.google.com/apis/api/people.googleapis.com
   - Google Tasks API: https://console.cloud.google.com/apis/api/tasks.googleapis.com
   - Google Sheets API: https://console.cloud.google.com/apis/api/sheets.googleapis.com
3. Configure OAuth consent screen: https://console.cloud.google.com/auth/branding
4. If your app is in "Testing", add test users: https://console.cloud.google.com/auth/audience
5. Create OAuth client:
   - Go to https://console.cloud.google.com/auth/clients
   - Click "Create Client"
   - Application type: "Desktop app"
   - Download the JSON file (usually named `client_secret_....apps.googleusercontent.com.json`)

### 2. Store Credentials

```bash
gog auth credentials ~/Downloads/client_secret_....json
```

### 3. Authorize Your Account

```bash
gog auth add you@gmail.com
```

This will open a browser window for OAuth authorization. The refresh token is stored securely in your system keychain.

### 4. Test Authentication

```bash
export GOG_ACCOUNT=you@gmail.com
gog gmail labels list
```

## Configuration

### Account Selection

Specify the account using either a flag or environment variable:

```bash
# Via flag
gog gmail search 'newer_than:7d' --account you@gmail.com

# Via environment
export GOG_ACCOUNT=you@gmail.com
gog gmail search 'newer_than:7d'
```

List configured accounts:

```bash
gog auth list
```

### Service Scopes

By default, `gog auth add` requests access to all services (gmail, calendar, drive, contacts, tasks, sheets, people). To request fewer scopes:

```bash
gog auth add you@gmail.com --services drive,calendar
```

If you need to add services later and Google doesn't return a refresh token, re-run with `--force-consent`:

```bash
gog auth add you@gmail.com --services all --force-consent
```

### Environment Variables

- `GOG_ACCOUNT` - Default account email to use (avoids repeating `--account` flag)
- `GOG_OUTPUT` - Output format: `text` (default) or `json`
- `GOG_COLOR` - Color mode: `auto` (default), `always`, or `never`
- `GOG_KEYRING_PASSWORD` - Password for encrypted on-disk keyring (Linux/WSL/container environments without OS keychain)

## Security

### Credential Storage

OAuth credentials are stored securely in your system's keychain:
- **macOS**: Keychain Access
- **Linux**: Secret Service (GNOME Keyring, KWallet)
- **Windows**: Credential Manager

The CLI uses [github.com/99designs/keyring](https://github.com/99designs/keyring) for secure storage.

If no OS keychain backend is available (e.g., Linux/WSL/container), keyring can fall back to an encrypted on-disk store and may prompt for a password; for non-interactive runs set `GOG_KEYRING_PASSWORD`.

### Best Practices

- **Never commit OAuth client credentials** to version control
- Store client credentials outside your project directory
- Use different OAuth clients for development and production
- Re-authorize with `--force-consent` if you suspect token compromise
- Remove unused accounts with `gog auth remove <email>`

## Commands

### Authentication

```bash
gog auth credentials <path>           # Store OAuth client credentials
gog auth add <email>                  # Authorize and store refresh token
gog auth list                         # List stored accounts
gog auth remove <email>               # Remove a stored refresh token
gog auth manage                       # Open accounts manager in browser
gog auth tokens                       # Manage stored refresh tokens
```

### Gmail

```bash
# Search and read
gog gmail search 'newer_than:7d' --max 10
gog gmail thread <threadId>
gog gmail thread <threadId> --download-attachments  # Download attachments to current dir
gog gmail get <messageId>
gog gmail get <messageId> --format metadata
gog gmail attachment <messageId> <attachmentId>
gog gmail attachment <messageId> <attachmentId> --out ./attachment.bin
gog gmail url <threadId>              # Print Gmail web URL

# Send and compose
gog gmail send --to a@b.com --subject "Hi" --body "Plain fallback"
gog gmail send --to a@b.com --subject "Hi" --body "Plain fallback" --body-html "<p>Hello</p>"
gog gmail drafts list
gog gmail drafts create --to a@b.com --subject "Draft"
gog gmail drafts send <draftId>

# Labels
gog gmail labels list
gog gmail labels get INBOX --output json  # Includes message counts
gog gmail labels create "My Label"
gog gmail labels update <labelId> --name "New Name"
gog gmail labels delete <labelId>

# Batch operations
gog gmail batch mark-read --query 'older_than:30d'
gog gmail batch delete --query 'from:spam@example.com'
gog gmail batch label --query 'from:boss@example.com' --add-labels IMPORTANT

# Filters
gog gmail filters list
gog gmail filters create --from 'noreply@example.com' --label 'Notifications'
gog gmail filters delete <filterId>

# Settings
gog gmail autoforward get
gog gmail autoforward enable --email forward@example.com
gog gmail autoforward disable
gog gmail forwarding list
gog gmail forwarding add --email forward@example.com
gog gmail sendas list
gog gmail sendas create --email alias@example.com
gog gmail vacation get
gog gmail vacation enable --subject "Out of office" --message "..."
gog gmail vacation disable

# Delegation (G Suite/Workspace)
gog gmail delegates list
gog gmail delegates add --email delegate@example.com
gog gmail delegates remove --email delegate@example.com

# Watch (Pub/Sub push)
gog gmail watch start --topic projects/<p>/topics/<t> --label INBOX
gog gmail watch serve --bind 127.0.0.1 --token <shared> --hook-url http://127.0.0.1:18789/hooks/agent
gog gmail watch serve --bind 0.0.0.0 --verify-oidc --oidc-email <svc@...> --hook-url <url>
gog gmail history --since <historyId>
```

Gmail watch (Pub/Sub push):
- Create Pub/Sub topic + push subscription (OIDC preferred; shared token ok for dev).
- Full flow + payload details: `docs/watch.md`.

### Calendar

```bash
# Calendars
gog calendar calendars
gog calendar acl <calendarId>         # List access control rules
gog calendar colors                   # List available event/calendar colors
gog calendar time --timezone America/New_York

# Events
gog calendar events <calendarId> --from 2025-01-01T00:00:00Z --to 2025-01-08T00:00:00Z --max 50
gog calendar events --all             # Fetch events from all calendars
gog calendar event <calendarId> <eventId>
gog calendar search "meeting" --from 2025-01-01T00:00:00Z --to 2025-01-31T00:00:00Z --max 50

# Create and update
gog calendar create <calendarId> \
  --summary "Meeting" \
  --start 2025-01-15T10:00:00Z \
  --end 2025-01-15T11:00:00Z

gog calendar create <calendarId> \
  --summary "Team Sync" \
  --start 2025-01-15T14:00:00Z \
  --end 2025-01-15T15:00:00Z \
  --organizer organizer@example.com \
  --color 5

gog calendar update <calendarId> <eventId> \
  --summary "Updated Meeting" \
  --start 2025-01-15T11:00:00Z \
  --end 2025-01-15T12:00:00Z

gog calendar delete <calendarId> <eventId>

# Invitations
gog calendar respond <calendarId> <eventId> --status accepted
gog calendar respond <calendarId> <eventId> --status declined
gog calendar respond <calendarId> <eventId> --status tentative

# Availability
gog calendar freebusy --calendars "primary,work@example.com" \
  --from 2025-01-15T00:00:00Z \
  --to 2025-01-16T00:00:00Z

gog calendar conflicts --calendars "primary,work@example.com" \
  --from 2025-01-15T00:00:00Z \
  --to 2025-01-22T00:00:00Z
```

### Drive

```bash
# List and search
gog drive ls --max 20
gog drive ls <folderId> --max 20      # List folder contents
gog drive search "invoice" --max 20
gog drive get <fileId>                # Get file metadata
gog drive url <fileId>                # Print Drive web URL

# Upload and download
gog drive upload ./path/to/file --folder <folderId>
gog drive download <fileId>

# Organize
gog drive mkdir "New Folder"
gog drive mkdir "New Folder" --parent <parentFolderId>
gog drive rename <fileId> "New Name"
gog drive move <fileId> --folder <destinationFolderId>
gog drive delete <fileId>             # Move to trash

# Permissions
gog drive permissions <fileId>
gog drive share <fileId> --email user@example.com --role reader
gog drive share <fileId> --email user@example.com --role writer
gog drive unshare <fileId> --permission-id <permissionId>
```

### Contacts

```bash
# Personal contacts
gog contacts list --max 50
gog contacts search "Ada" --max 50
gog contacts get people/<resourceName>
gog contacts get user@example.com     # Get by email

# Other contacts (people you've interacted with)
gog contacts other list --max 50
gog contacts other search "John" --max 50

# Create and update
gog contacts create \
  --given-name "John" \
  --family-name "Doe" \
  --email "john@example.com" \
  --phone "+1234567890"

gog contacts update people/<resourceName> \
  --given-name "Jane" \
  --email "jane@example.com"

gog contacts delete people/<resourceName>

# Workspace directory (requires Google Workspace)
gog contacts directory list --max 50
gog contacts directory search "Jane" --max 50
```

### Tasks

```bash
# Task lists
gog tasks lists --max 50
gog tasks lists create <title>

# Tasks in a list
gog tasks list <tasklistId> --max 50
gog tasks add <tasklistId> --title "Task title"
gog tasks update <tasklistId> <taskId> --title "New title"
gog tasks done <tasklistId> <taskId>
gog tasks undo <tasklistId> <taskId>
gog tasks delete <tasklistId> <taskId>
gog tasks clear <tasklistId>
```

### Sheets

```bash
# Read
gog sheets metadata <spreadsheetId>
gog sheets get <spreadsheetId> 'Sheet1!A1:B10'

# Write
gog sheets update <spreadsheetId> 'A1' 'val1|val2,val3|val4'
gog sheets update <spreadsheetId> 'A1' --json '[["a","b"],["c","d"]]'
gog sheets append <spreadsheetId> 'Sheet1!A:C' 'new|row|data'
gog sheets clear <spreadsheetId> 'Sheet1!A1:B10'

# Create
gog sheets create "My New Spreadsheet" --sheets "Sheet1,Sheet2"
```

### People

```bash
# Profile
gog people me
```

## Output Formats

### Text

Human-readable output with colors (default):

```bash
$ gog gmail search 'newer_than:7d' --max 3
THREAD_ID           SUBJECT                           FROM                  DATE
18f1a2b3c4d5e6f7    Meeting notes                     alice@example.com     2025-01-10
17e1d2c3b4a5f6e7    Invoice #12345                    billing@vendor.com    2025-01-09
16d1c2b3a4e5f6d7    Project update                    bob@example.com       2025-01-08
```

### JSON

Machine-readable output for scripting and automation:

```bash
$ gog gmail search 'newer_than:7d' --max 3 --output json
{
  "threads": [
    {
      "id": "18f1a2b3c4d5e6f7",
      "snippet": "Meeting notes from today...",
      "messages": [...]
    },
    ...
  ]
}
```

Data goes to stdout, errors and progress to stderr for clean piping:

```bash
gog --output json drive ls --max 5 | jq '.files[] | select(.mimeType=="application/pdf")'
```

Useful pattern:

```bash
gog --output json ... | jq .
```

## Examples

### Search recent emails and download attachments

```bash
# Search for emails from the last week
gog gmail search 'newer_than:7d has:attachment' --max 10

# Get thread details and download attachments
gog gmail thread <threadId> --download-attachments
```

### Create a calendar event with attendees

```bash
# Find a free time slot
gog calendar freebusy --calendars "primary" \
  --from 2025-01-15T00:00:00Z \
  --to 2025-01-16T00:00:00Z

# Create the meeting
gog calendar create primary \
  --summary "Team Standup" \
  --start 2025-01-15T10:00:00Z \
  --end 2025-01-15T10:30:00Z \
  --attendees "alice@example.com,bob@example.com"
```

### Find and download files from Drive

```bash
# Search for PDFs
gog drive search "invoice filetype:pdf" --max 20 --output json | \
  jq -r '.files[] | .id' | \
  while read fileId; do
    gog drive download "$fileId"
  done
```

### Manage multiple accounts

```bash
# Check personal Gmail
gog gmail search 'is:unread' --account personal@gmail.com

# Check work Gmail
gog gmail search 'is:unread' --account work@company.com

# Or set default
export GOG_ACCOUNT=work@company.com
gog gmail search 'is:unread'
```

### Update a Google Sheet from a CSV

```bash
# Convert CSV to pipe-delimited format and update sheet
cat data.csv | tr ',' '|' | \
  gog sheets update <spreadsheetId> 'Sheet1!A1'
```

### Batch process Gmail threads

```bash
# Mark all emails from a sender as read
gog gmail batch mark-read --query 'from:noreply@example.com'

# Archive old emails
gog gmail batch archive --query 'older_than:1y'

# Label important emails
gog gmail batch label --query 'from:boss@example.com' --add-labels IMPORTANT
```

## Advanced Features

### Debug Mode

Enable verbose output for troubleshooting:

```bash
gog --debug gmail search 'newer_than:7d'
# Shows API requests and responses
```

## Global Flags

All commands support these flags:

- `--account <email>` - Account to use (overrides GOG_ACCOUNT)
- `--output <format>` - Output format: `text` or `json` (default: text)
- `--color <mode>` - Color mode: `auto`, `always`, or `never` (default: auto)
- `--debug` - Enable debug logging
- `--help` - Show help for any command

## Shell Completions

Generate shell completions for your preferred shell:

### Bash

```bash
# macOS (with Homebrew)
gog completion bash > $(brew --prefix)/etc/bash_completion.d/gog

# Linux
gog completion bash > /etc/bash_completion.d/gog

# Or load directly in your current session
source <(gog completion bash)
```

### Zsh

```zsh
# Generate completion file
gog completion zsh > "${fpath[1]}/_gog"

# Or add to .zshrc for automatic loading
echo 'eval "$(gog completion zsh)"' >> ~/.zshrc

# Enable completions if not already enabled
echo "autoload -U compinit; compinit" >> ~/.zshrc
```

### Fish

```fish
gog completion fish > ~/.config/fish/completions/gog.fish
```

### PowerShell

```powershell
# Load for current session
gog completion powershell | Out-String | Invoke-Expression

# Or add to profile for all sessions
gog completion powershell >> $PROFILE
```

After installing completions, start a new shell session for changes to take effect.

## Development

After cloning, install git hooks:

```bash
make setup
```

This installs [lefthook](https://github.com/evilmartians/lefthook) pre-commit and pre-push hooks for linting and testing.

Pinned tools (installed into `.tools/`):

- Format: `make fmt` (goimports + gofumpt)
- Lint: `make lint` (golangci-lint)
- Test: `make test`

CI runs format checks, tests, and lint on push/PR.

### pnpm Shortcut

Build and run in one step:

```bash
pnpm gog auth add you@gmail.com
```

For clean stdout when scripting:

```bash
pnpm -s gog --output json gmail search "from:me" | jq .
```

## License

MIT

## Links

- [GitHub Repository](https://github.com/steipete/gogcli)
- [Gmail API Documentation](https://developers.google.com/gmail/api)
- [Google Calendar API Documentation](https://developers.google.com/calendar)
- [Google Drive API Documentation](https://developers.google.com/drive)
- [Google People API Documentation](https://developers.google.com/people)
- [Google Tasks API Documentation](https://developers.google.com/tasks)
- [Google Sheets API Documentation](https://developers.google.com/sheets)

## Credits

This project is inspired by Mario Zechner's original CLIs:

- [gmcli](https://github.com/badlogic/gmcli)
- [gccli](https://github.com/badlogic/gccli)
- [gdcli](https://github.com/badlogic/gdcli)
