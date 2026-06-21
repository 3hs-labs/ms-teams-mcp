# Microsoft Teams & Outlook MCP Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that provides access to Microsoft Teams chats/channels, Outlook emails, calendar events, and SharePoint files via the Microsoft Graph API.

Use natural language in Claude Code, VS Code, Claude Desktop, or any MCP client to read Teams messages, search emails, manage calendar events, browse files, and more.

## Features

- **Teams** — List teams/channels, read/send/reply channel messages, create/read/send 1:1 & group chats
- **Outlook** — List/read/search/send/reply/forward emails, list mail folders
- **Calendar** — List/create/update/delete events, recurring events, reminders
- **SharePoint** — List channel files, read file contents (text/xlsx), build & search file index
- **People** — Search users across the organization
- **Utilities** — Unread summary, auth status, version update check
- **Auth** — Device Code Flow authentication directly from MCP or CLI
- **Web Transport** — SSE and streamable-http support for remote access from Claude Desktop, Claude web, and other MCP clients

## Prerequisites

### 1. Register an Azure AD App

1. Go to [Azure Portal](https://portal.azure.com/) > **Azure Active Directory** > **App registrations** > **New registration**
2. Set **Redirect URI** to `https://login.microsoftonline.com/common/oauth2/nativeclient` (Mobile and desktop applications)
3. Under **Certificates & secrets**, create a client secret
4. Under **API permissions**, add these Microsoft Graph **delegated permissions**:

| Permission | Purpose |
|------------|---------|
| `User.Read` | User profile |
| `Mail.Read` | Read emails |
| `Mail.Send` | Send, reply, forward emails |
| `Chat.Read` | Read chats |
| `Chat.ReadWrite` | Send chat messages, create chats |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read channel messages |
| `ChannelMessage.Send` | Send channel messages |
| `Team.ReadBasic.All` | List teams |
| `Files.Read.All` | Read SharePoint files |
| `People.Read` | Search users |
| `Calendars.ReadWrite` | Read/create/update/delete calendar events |

5. Click **Grant admin consent**

### 2. Note Your Credentials

You will need these three values:

- `MS_CLIENT_ID` — Application (client) ID
- `MS_CLIENT_SECRET` — Client secret value
- `MS_TENANT_ID` — Directory (tenant) ID

### 3. Install uv

This project uses [`uvx`](https://docs.astral.sh/uv/) to run the MCP server:

**Linux / macOS:**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

**Windows (PowerShell):**
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Restart your terminal, then verify: `uvx --version`

## Installation & Usage

### Claude Code (Recommended)

No pre-installation needed — `uvx` automatically downloads and runs the package. Just register and authenticate from within Claude.

**Linux / macOS:**
```bash
claude mcp add ms-teams \
  -s user \
  -e MS_CLIENT_ID=<your-client-id> \
  -e MS_CLIENT_SECRET=<your-client-secret> \
  -e MS_TENANT_ID=<your-tenant-id> \
  -- uvx --from ms-teams-mcp ms-teams-mcp
```

**Windows (PowerShell):**
```powershell
claude mcp add ms-teams `
  -s user `
  -e MS_CLIENT_ID=<your-client-id> `
  -e MS_CLIENT_SECRET=<your-client-secret> `
  -e MS_TENANT_ID=<your-tenant-id> `
  -- uvx --from ms-teams-mcp ms-teams-mcp
```

> **Windows Note**: If `uvx` is not found, use `uvx.cmd` or the full path `%USERPROFILE%\.local\bin\uvx.cmd`.

On first use, ask Claude to call the `authenticate` tool — it will guide you through Device Code Flow login in your browser.

### VS Code

Add to `.vscode/mcp.json` in your project root (or under `"mcp"` key in user `settings.json` for global access):

```json
{
  "servers": {
    "ms-teams": {
      "command": "uvx",
      "args": [
        "--from",
        "ms-teams-mcp",
        "ms-teams-mcp"
      ],
      "env": {
        "MS_CLIENT_ID": "<your-client-id>",
        "MS_CLIENT_SECRET": "<your-client-secret>",
        "MS_TENANT_ID": "<your-tenant-id>"
      }
    }
  }
}
```

> **Windows**: Use `"command": "uvx.cmd"` if `uvx` is not found.

### Claude Desktop

Add to your config file:

- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "ms-teams": {
      "command": "uvx",
      "args": [
        "--from",
        "ms-teams-mcp",
        "ms-teams-mcp"
      ],
      "env": {
        "MS_CLIENT_ID": "<your-client-id>",
        "MS_CLIENT_SECRET": "<your-client-secret>",
        "MS_TENANT_ID": "<your-tenant-id>"
      }
    }
  }
}
```

> **Windows**: Replace `"command": "uvx"` with `"command": "uvx.cmd"` if `uvx` is not found.

### pip Install

```bash
pip install ms-teams-mcp

claude mcp add ms-teams \
  -s user \
  -e MS_CLIENT_ID=<your-client-id> \
  -e MS_CLIENT_SECRET=<your-client-secret> \
  -e MS_TENANT_ID=<your-tenant-id> \
  -- ms-teams-mcp
```

### Authentication

On first use, call the `authenticate` tool from your MCP client — it will provide a Device Code Flow URL to sign in via your browser.

Alternatively, you can authenticate from the CLI before starting:

```bash
# Using uvx
uvx --from ms-teams-mcp \
  ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>

# Using pip install
ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>
```

Token is cached at `~/.ms_mcp_token.json` and silently renewed on subsequent runs.

### Upgrade

```bash
# uvx — clear cache, auto-downloads latest on next run
uv cache clean

# pip — force reinstall
pip install --upgrade --force-reinstall ms-teams-mcp
```

The server automatically checks for updates on startup (once per 24h) and notifies via stderr. You can also call the `check_update` tool from any MCP client.

### Verify & Remove

```bash
claude mcp list                       # List registered MCP servers
claude mcp remove ms-teams     # Remove server
```

## Available MCP Tools (32)

| Category | Tool | Description |
|----------|------|-------------|
| **Auth** | `auth_status` | Check authentication status and token validity |
| | `authenticate` | Authenticate via Device Code Flow |
| **Teams** | `list_teams` | List joined teams |
| | `list_channels` | List channels in a team |
| | `list_channel_messages` | Read channel messages (max 50) |
| | `send_channel_message` | Send a message to a channel |
| | `reply_to_channel_message` | Reply to a channel message |
| **Chats** | `list_chats` | List 1:1 and group chats (max 50) |
| | `list_chat_messages` | Read chat messages (max 50) |
| | `send_chat_message` | Send a chat message |
| | `reply_to_chat_message` | Reply to a chat message |
| | `create_chat` | Create a new 1:1 or group chat |
| **Outlook** | `list_emails` | List emails (max 1000) |
| | `read_email` | Read full email body |
| | `search_emails` | Search emails (max 1000) |
| | `send_email` | Send an email |
| | `reply_email` | Reply or reply-all to an email |
| | `forward_email` | Forward an email |
| | `list_mail_folders` | List mail folders |
| **Calendar** | `list_calendar_events` | List calendar events in a date range |
| | `create_calendar_event` | Create a new calendar event |
| | `update_calendar_event` | Update an existing calendar event |
| | `delete_calendar_event` | Delete a calendar event |
| | `create_recurring_event` | Create a recurring event (daily/weekly/monthly/yearly) |
| | `create_reminder` | Create a reminder with alert |
| **Files** | `list_channel_files` | List files in a channel (max 200) |
| | `read_channel_file` | Read file contents (text/xlsx, 5MB limit) |
| | `build_file_index` | Build searchable index of all accessible files |
| | `search_file_index` | Search the file index by keyword |
| **People** | `search_users` | Search users in the organization |
| **Utilities** | `get_unread_summary` | Summarize unread emails and chats |
| | `check_update` | Check for newer version on GitHub |

All list tools support `top`, `skip`, and `next_link` parameters for pagination.

### Example Prompts

**Teams — Channels**
```
> Show my Teams list
> Show channels in the "Engineering" team
> Show the last 10 messages in the "General" channel of "Project A" team
> Send "Build completed successfully" to the #releases channel
> Reply to the latest message in #general with "Thanks for the update!"
```

**Teams — Chats**
```
> Show my recent chats
> Show messages from my chat with Sarah
> Send "Are you available for a quick call?" to my chat with David
> Create a group chat with john@example.com and jane@example.com about "Q2 Planning"
```

**Outlook — Email**
```
> Check my inbox for today's emails
> Show my unread emails
> Read the latest email from my manager
> Search emails for "quarterly report" from last week
> Send an email to john@example.com with subject "Project Update" and summarize today's progress
> Reply to the latest email from Sarah saying "I'll review it by EOD"
> Forward the budget email to the finance team at finance@example.com
> Show my mail folders
```

**Calendar**
```
> Show my calendar for this week
> What meetings do I have tomorrow?
> Create a meeting with david@example.com tomorrow at 2pm for 1 hour about "Design Review"
> Schedule a Teams online meeting with the frontend team next Monday 10am-11am
> Create a weekly recurring standup every weekday at 9:30am starting next Monday
> Set up a monthly team sync on the first Tuesday of each month
> Remind me about the report deadline at 3pm tomorrow
> Move my 2pm meeting to 4pm
> Cancel the design review meeting on Friday
```

**Files**
```
> Show files in the "General" channel of "Project A" team
> Read the project-plan.xlsx file from the Documents channel
> Build an index of all my accessible files
> Search my file index for "budget"
```

**People & Utilities**
```
> Search for a user named "Kim" in the organization
> Find the email address of someone in the marketing team
> Summarize my unread emails and chat messages
> Check if there's a newer version of the MCP server
```

**Combined Workflows**
```
> Check my unread messages and emails, then give me a morning briefing
> Find all emails about "Project X", summarize them, and send a status update to the team channel
> Look at my calendar for next week and find a free slot for a 1-hour meeting
> Read the latest messages in #engineering and reply with a summary of today's deploy
> Search for "budget" in my emails and files, then compile the key numbers
```

<details>
<summary><strong>한국어 사용 예시</strong></summary>

**Teams — 채널**
```
> 내 팀 목록 보여줘
> "개발팀" 팀의 채널 목록 보여줘
> "프로젝트A" 팀의 "일반" 채널 최근 메시지 10개 보여줘
> #releases 채널에 "빌드 완료되었습니다" 메시지 보내줘
> #general 최신 메시지에 "확인했습니다, 감사합니다!" 답장해줘
```

**Teams — 채팅**
```
> 최근 채팅 목록 보여줘
> 김민수님과의 채팅 내용 보여줘
> 이영희님에게 "잠깐 통화 가능하신가요?" 메시지 보내줘
> john@example.com, jane@example.com과 "2분기 기획" 주제로 그룹 채팅 만들어줘
```

**Outlook — 메일**
```
> 오늘 받은 메일 확인해줘
> 안 읽은 메일 보여줘
> 팀장님한테 온 최신 메일 읽어줘
> 지난주 "분기 보고서" 관련 메일 검색해줘
> john@example.com에게 "프로젝트 현황" 제목으로 오늘 진행 상황 정리해서 메일 보내줘
> 김과장님 메일에 "오늘 중으로 검토하겠습니다" 답장해줘
> 예산 관련 메일을 재무팀 finance@example.com으로 전달해줘
> 메일 폴더 목록 보여줘
```

**캘린더**
```
> 이번 주 일정 보여줘
> 내일 회의 뭐 있어?
> 내일 오후 2시에 david@example.com과 1시간 "디자인 리뷰" 회의 잡아줘
> 다음 주 월요일 오전 10시~11시 프론트엔드 팀 Teams 온라인 회의 만들어줘
> 다음 주 월요일부터 매일 오전 9시 30분 스탠드업 반복 일정 만들어줘
> 매월 첫째 화요일에 팀 싱크 회의 설정해줘
> 내일 오후 3시에 보고서 마감 리마인더 설정해줘
> 오후 2시 회의를 4시로 변경해줘
> 금요일 디자인 리뷰 회의 취소해줘
```

**파일**
```
> "프로젝트A" 팀 "일반" 채널의 파일 목록 보여줘
> Documents 채널의 project-plan.xlsx 파일 읽어줘
> 내가 접근 가능한 모든 파일 인덱스 만들어줘
> 파일 인덱스에서 "예산" 검색해줘
```

**사용자 & 유틸리티**
```
> 조직에서 "김" 이름 가진 사용자 검색해줘
> 마케팅팀 누군가의 이메일 주소 찾아줘
> 안 읽은 메일이랑 채팅 메시지 요약해줘
> MCP 서버 새 버전 있는지 확인해줘
```

**복합 시나리오**
```
> 안 읽은 메시지랑 메일 확인해서 오늘 아침 브리핑 해줘
> "프로젝트X" 관련 메일 전부 찾아서 요약하고, 팀 채널에 현황 공유해줘
> 다음 주 캘린더 확인하고 1시간짜리 회의 가능한 시간 찾아줘
> #engineering 최신 메시지 읽고 오늘 배포 내용 요약해서 답장해줘
> 메일이랑 파일에서 "예산" 검색해서 주요 수치 정리해줘
```

</details>

## Web Transport (SSE / Streamable HTTP)

In addition to the default `stdio` mode, the server can run as an HTTP service — enabling remote connections from Claude Desktop, Claude web (claude.ai), and other MCP clients.

### Quick Start

```bash
# Start with streamable-http (default web transport)
ms-teams-mcp serve

# Start with SSE transport
ms-teams-mcp serve --transport sse

# Custom host/port (default: 127.0.0.1:7979)
ms-teams-mcp serve --host 0.0.0.0 --port 9000
```

### Claude Desktop (SSE)

Update your Claude Desktop config to connect via URL instead of launching a subprocess:

- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "ms-teams": {
      "url": "http://localhost:7979/sse"
    }
  }
}
```

Then start the server separately:

```bash
MS_CLIENT_ID=<id> MS_CLIENT_SECRET=<secret> MS_TENANT_ID=<tenant> \
  ms-teams-mcp serve --transport sse
```

### Claude Web (claude.ai)

Claude web requires **streamable-http** over **HTTPS** on a public URL. You need a Pro, Team, or Enterprise plan.

**1. Start the server:**

```bash
ms-teams-mcp serve --host 0.0.0.0 --port 7979
```

**2. Set up a reverse proxy with HTTPS** (e.g., nginx + Let's Encrypt):

```nginx
server {
    listen 443 ssl;
    server_name mcp.your-domain.com;

    ssl_certificate     /etc/letsencrypt/live/mcp.your-domain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/mcp.your-domain.com/privkey.pem;

    location / {
        proxy_pass http://127.0.0.1:7979;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
    }
}
```

**3. Connect from Claude web:**

Go to Claude.ai → Settings → MCP Connectors → Add → enter your server URL:

```
https://mcp.your-domain.com/mcp
```

### Transport Comparison

| Mode | Command | Endpoint | Use Case |
|------|---------|----------|----------|
| stdio | `ms-teams-mcp` | — | Claude Code, VS Code (local) |
| SSE | `ms-teams-mcp serve --transport sse` | `http://host:7979/sse` | Claude Desktop (local/remote) |
| streamable-http | `ms-teams-mcp serve` | `http://host:7979/mcp` | Claude web, remote clients |

> **Note:** The server uses a single token cache (`~/.ms_mcp_token.json`), so all clients connecting to the same server share one Microsoft account. For per-user access, each user should run their own server instance.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `uvx` not found (Windows) | Use `uvx.cmd` or full path `%USERPROFILE%\.local\bin\uvx.cmd` |
| `uvx` not found (Linux) | Run `source ~/.bashrc` or restart terminal after uv install |
| Token expired | Re-run `ms-teams-mcp auth ...` or call `authenticate` tool |
| Permission denied (403) | Check Azure AD app has required permissions with admin consent |
| Rate limit (429) | Wait and retry — Graph API throttles excessive requests |
| `cp949` decode error | File encoding not supported; only UTF-8 and CP949 are handled |
| Server not starting | Verify env vars `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_TENANT_ID` are set |

## Project Structure

```
├── ms_teams_mcp/
│   ├── __init__.py    # Package init
│   └── server.py      # MCP server (single-file architecture)
├── pyproject.toml     # Package config & dependencies
├── CLAUDE.md          # Claude Code instructions
└── README.md
```

## License

MIT
