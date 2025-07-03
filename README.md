# MCP Outlook Tools

A Model Context Protocol (MCP) server implementation that enables AI assistants to interact with Microsoft Outlook for calendar management, email operations, and search functionality.

## Features

- 📅 **Calendar Management**
  - Get calendar items within a date range
  - Add new appointments with full details
  - Support for categories and busy status

- 📧 **Email Operations**
  - Send emails with To/CC recipients
  - Display confirmation before sending
  - Full body formatting support

- 🔍 **Email Search**
  - Search emails by date and keyword
  - Extract user information from addresses
  - Japanese text encoding support

## Requirements

- Windows OS (required for pywin32)
- Microsoft Outlook installed and configured
- Python 3.10 or higher
- MCP-compatible AI assistant (e.g., Claude Desktop)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/wmoto-ai/mcp-outlook-tools.git
cd mcp-outlook-tools
```

2. Install dependencies using uv:
```bash
uv pip install -e .
```

Or using pip:
```bash
pip install -e .
```

## Configuration

### For Claude Desktop

Add the following to your Claude Desktop configuration file:

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "outlook-tools": {
      "command": "uv",
      "args": [
        "--directory",
        "C:/path/to/mcp-outlook-tools",
        "run",
        "--with-editable",
        ".",
        "-m",
        "outlook_tools.server"
      ],
      "cwd": "C:/path/to/mcp-outlook-tools",
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

## Usage

Once configured, the following tools are available in your AI assistant:

### `add_appointment`
```
Add a new appointment to Outlook calendar
Parameters:
- subject: Appointment title
- start_time: Start datetime (YYYY-MM-DD HH:MM)
- end_time: End datetime (YYYY-MM-DD HH:MM)
- location: Meeting location (optional)
- description: Detailed description (optional)
- categories: Comma-separated categories (optional)
- busy_status: 0=Free, 1=Tentative, 2=Busy, 3=Out of Office (default: 1)
```

### `get_calendar`
```
Get calendar items for a date range
Parameters:
- start_date: Start date (YYYY-MM-DD)
- end_date: End date (YYYY-MM-DD)
```

### `send_email`
```
Send an email via Outlook
Parameters:
- to: Recipient email addresses (semicolon-separated)
- cc: CC recipients (semicolon-separated)
- subject: Email subject
- body: Email body text
```

## Project Structure

```
mcp-outlook-tools/
├── src/
│   └── outlook_tools/
│       ├── __init__.py
│       ├── server.py           # MCP server implementation
│       ├── calendar_service.py # Calendar operations
│       └── search_service.py   # Email search operations
├── test/                       # Test files
├── pyproject.toml             # Project configuration
└── README.md                  # This file
```

## Development

### Running Tests
```bash
pytest test/
```

### Type Checking
```bash
pyright src/
```

### Linting
```bash
ruff check src/
```

## Security Notes

- This tool requires access to your local Outlook installation
- Emails are displayed before sending for user confirmation
- No credentials are stored in the code
- All operations use Windows COM interface with existing Outlook authentication

## Limitations

- Windows only (due to pywin32 dependency)
- Requires Outlook to be installed and configured
- Time zone handling assumes JST (+9 hours)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with [FastMCP](https://github.com/modelcontextprotocol/fastmcp) framework
- Uses pywin32 for Outlook COM interface