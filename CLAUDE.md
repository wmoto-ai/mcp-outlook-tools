# MCP Outlook Tools Project Overview

## Project Summary
This is an MCP (Model Context Protocol) server implementation that provides Outlook integration capabilities for AI assistants. The project enables AI models to interact with Microsoft Outlook for calendar management, email sending, and email searching functionalities.

## Technical Stack
- **Language**: Python 3.10+
- **Protocol**: MCP (Model Context Protocol) v1.2.0+
- **Dependencies**: 
  - pywin32 (for Windows COM interface)
  - python-dateutil (for date parsing)
  - FastMCP (MCP server framework)

## Core Features

### 1. Calendar Management (`calendar_service.py`)
- **Get Calendar Items**: Retrieve appointments within a specified date range
- **Add Appointments**: Create new calendar appointments with:
  - Subject, start/end times
  - Location and description
  - Categories and busy status

### 2. Email Operations (`server.py`)
- **Send Email**: Create and send emails with:
  - To, CC recipients
  - Subject and body
  - Display before sending (for user confirmation)

### 3. Email Search (`search_service.py`)
- **Search Emails**: Find emails by date and keyword
- **User Information**: Extract user details from email addresses
- **Encoding Support**: Handle Japanese text encoding issues

## Project Structure
```
mcp-outlook-tools/
├── src/
│   └── outlook_tools/           # Main package
│       ├── __init__.py
│       ├── server.py           # MCP server and tool definitions
│       ├── calendar_service.py # Calendar operations
│       └── search_service.py   # Email search operations
├── test/                       # Test files
│   ├── retrieve_appointments.py
│   └── test_calendar_service.py
├── pyproject.toml             # Project configuration
└── uv.lock                    # Dependency lock file
```

## MCP Tools Exposed

### `add_appointment`
Adds a new appointment to Outlook calendar.
- Parameters: subject, start_time, end_time, location, description, categories, busy_status
- Returns: Success/failure message

### `get_calendar`
Retrieves calendar items for a specified date range.
- Parameters: start_date, end_date (YYYY-MM-DD format)
- Returns: List of appointments with details

### `send_email`
Sends an email through Outlook with display confirmation.
- Parameters: to, cc, subject, body
- Returns: Success/failure message

## Platform Requirements
- Windows OS (required for pywin32)
- Microsoft Outlook installed and configured
- Python 3.10 or higher

## Development Tools
- pyright: Type checking
- pytest: Testing framework
- ruff: Linting

## Key Implementation Details
- Uses Windows COM interface to interact with Outlook
- Handles datetime parsing with timezone adjustments (JST +9 hours)
- Includes encoding fixes for Japanese text
- FastMCP framework for easy MCP server implementation

## Security Considerations
- Requires local Outlook installation with proper authentication
- Email display before sending provides user confirmation
- No credentials are stored in the code

## Future Enhancements
- Additional email search filters
- Contact management features
- Task and note management
- Cross-platform support (beyond Windows)