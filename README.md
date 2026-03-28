# Microsoft Teams MCP Server

A Model Context Protocol (MCP) server for Microsoft Teams integration with OAuth 2.0 authentication.

## Features

- **OAuth 2.0 Authentication**: Secure Microsoft OAuth flow with automatic token refresh
- **MCP Protocol Support**: Full implementation of the Model Context Protocol
- **Teams Operations**:
  - Search chats and channel messages
  - Fetch specific messages by path
  - Get chat members
  - Get user profile
- **Session Management**: Persistent sessions with automatic token refresh
- **TypeScript**: Fully typed with TypeScript for better development experience
- **Cloudflare Workers**: Deployable to Cloudflare Workers for global edge performance

## Prerequisites

- Node.js 18+ 
- Microsoft Azure App Registration
- OAuth 2.0 credentials (Client ID and Client Secret)

## Setup

### 1. Microsoft Azure Configuration

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure your application:
   - Name: `Teams MCP Server`
   - Supported account types: Choose based on your needs
   - Redirect URI: `Web` - `http://localhost:3003/oauth/callback`
5. After creation, note the **Application (client) ID** and **Directory (tenant) ID**
6. Go to **Certificates & secrets** > **New client secret**
7. Copy the client secret value
8. Go to **API permissions** > **Add a permission** > **Microsoft Graph**
9. Add the following **Delegated permissions**:
   - `Chat.Read`
   - `ChannelMessage.Read.All`
   - `User.Read`
10. Click **Grant admin consent** (if you have admin privileges)

### 2. Installation

```bash
# Clone or navigate to the project directory
cd teams-mcp-server

# Install dependencies
npm install
```

### 3. Configuration

Create a `.env` file in the root directory:

```bash
cp env.example .env
```

Edit `.env` with your configuration:

```env
# Server Configuration
PORT=3003
NODE_ENV=development

# Microsoft OAuth Configuration
MICROSOFT_CLIENT_ID=your_microsoft_client_id_here
MICROSOFT_CLIENT_SECRET=your_microsoft_client_secret_here
MICROSOFT_REDIRECT_URI=http://localhost:3003/oauth/callback
MICROSOFT_TENANT_ID=common

# Frontend URL (where to redirect after OAuth)
FRONTEND_URL=http://localhost:5173

# MCP Server Configuration
MCP_SERVER_NAME=teams-mcp-server
MCP_SERVER_VERSION=1.0.0
```

**Tenant ID Options:**
- `common` - Multi-tenant (personal and work/school accounts)
- `organizations` - Work/school accounts only
- `consumers` - Personal Microsoft accounts only
- `{tenant-id}` - Specific tenant/organization

## Usage

### Development Mode

```bash
npm run dev
```

### Production Mode

```bash
# Build
npm run build

# Start
npm start
```

### Cloudflare Workers Deployment

```bash
# Create KV namespace for sessions
npm run kv:create

# Update wrangler.toml with your KV namespace ID

# Set secrets
wrangler secret put MICROSOFT_CLIENT_ID
wrangler secret put MICROSOFT_CLIENT_SECRET
wrangler secret put MICROSOFT_REDIRECT_URI
wrangler secret put MICROSOFT_TENANT_ID
wrangler secret put FRONTEND_URL

# Deploy
npm run deploy
```

## API Endpoints

### Health Check
```
GET /health
```

### OAuth Flow

#### 1. Initiate Authorization
```
GET /oauth/authorize?userId={userId}&state={state}
```

Response:
```json
{
  "authorizationUrl": "https://login.microsoftonline.com/...",
  "state": "your-state-value"
}
```

#### 2. OAuth Callback (handled automatically)
```
GET /oauth/callback?code={code}&state={state}
```

#### 3. Refresh Token
```
POST /oauth/refresh
Content-Type: application/json

{
  "userId": "user123",
  "refreshToken": "refresh_token_here"
}
```

#### 4. Disconnect
```
POST /oauth/disconnect
Content-Type: application/json

{
  "userId": "user123"
}
```

### MCP Endpoint

```
POST /mcp
GET /mcp
DELETE /mcp
```

The MCP endpoint follows the Model Context Protocol specification for tool execution.

## MCP Tools

### 1. teams_search

Search Microsoft Teams chats and channel messages.

**Scopes Required:** `Chat.Read`, `ChannelMessage.Read.All`

**Input:**
```json
{
  "userId": "user123",
  "query": "meeting notes",
  "chatId": "optional_chat_id",
  "channelId": "optional_channel_id",
  "teamId": "optional_team_id",
  "maxResults": 20,
  "fromDate": "2024-01-01T00:00:00Z",
  "toDate": "2024-12-31T23:59:59Z"
}
```

**Output:**
```json
{
  "messages": [
    {
      "id": "msg123",
      "chatId": "chat123",
      "from": {
        "displayName": "John Doe",
        "id": "user456"
      },
      "body": {
        "content": "Here are the meeting notes...",
        "contentType": "text"
      },
      "createdDateTime": "2024-01-15T10:30:00Z",
      "subject": "Meeting Notes",
      "webUrl": "https://teams.microsoft.com/..."
    }
  ],
  "count": 1
}
```

### 2. teams_fetch

Fetch a specific Teams message by path.

**Scopes Required:** `Chat.Read`, `ChannelMessage.Read.All`

**Input:**
```json
{
  "userId": "user123",
  "path": "chats/{chatId}/messages/{messageId}"
}
```

Or for channel messages:
```json
{
  "userId": "user123",
  "path": "teams/{teamId}/channels/{channelId}/messages/{messageId}"
}
```

**Output:**
```json
{
  "id": "msg123",
  "chatId": "chat123",
  "from": {
    "displayName": "John Doe",
    "id": "user456"
  },
  "body": {
    "content": "Message content here",
    "contentType": "text"
  },
  "createdDateTime": "2024-01-15T10:30:00Z",
  "webUrl": "https://teams.microsoft.com/..."
}
```

### 3. teams_get_chat_members

List the members of a Microsoft Teams chat.

**Scopes Required:** `Chat.Read`

**Input:**
```json
{
  "userId": "user123",
  "chatId": "chat123"
}
```

**Output:**
```json
{
  "members": [
    {
      "id": "member123",
      "displayName": "John Doe",
      "userId": "user456",
      "email": "john.doe@company.com",
      "roles": ["owner"]
    }
  ],
  "count": 1
}
```

### 4. teams_get_profile

Return the authenticated Teams user's profile.

**Scopes Required:** `User.Read`

**Input:**
```json
{
  "userId": "user123"
}
```

**Output:**
```json
{
  "id": "user123",
  "displayName": "John Doe",
  "givenName": "John",
  "surname": "Doe",
  "userPrincipalName": "john.doe@company.com",
  "mail": "john.doe@company.com",
  "jobTitle": "Software Engineer",
  "officeLocation": "Building 1",
  "mobilePhone": "+1234567890",
  "businessPhones": ["+1234567890"]
}
```

## Architecture

```
┌─────────────────┐
│   Frontend      │
│   (ZeroTwo)     │
└────────┬────────┘
         │
         │ OAuth Flow
         ▼
┌─────────────────┐      ┌──────────────────┐
│  Teams MCP      │◄────►│  Microsoft OAuth │
│  Server         │      │  & Graph API     │
└────────┬────────┘      └──────────────────┘
         │
         │ MCP Protocol
         ▼
┌─────────────────┐
│   Backend       │
│   (ZeroTwoApi)  │
└─────────────────┘
```

## Development

### Project Structure

```
teams-mcp-server/
├── src/
│   ├── auth/
│   │   └── oauth-manager.ts    # OAuth token management
│   ├── config/
│   │   └── index.ts            # Configuration management
│   ├── teams/
│   │   └── client.ts           # Teams/Graph API wrapper
│   ├── mcp/
│   │   ├── server.ts           # MCP server implementation
│   │   └── tools.ts            # MCP tool definitions
│   ├── types/
│   │   └── index.ts            # TypeScript type definitions
│   ├── utils/
│   │   └── logger.ts           # Logging utility
│   ├── index.ts                # Entry point
│   └── worker.ts               # Cloudflare Worker entry
├── package.json
├── tsconfig.json
├── wrangler.toml
├── env.example
└── README.md
```

### Scripts

- `npm run dev` - Start development server with hot reload
- `npm run dev:worker` - Start Cloudflare Workers development server
- `npm run build` - Build TypeScript to JavaScript
- `npm run build:worker` - Build for Cloudflare Workers (dry run)
- `npm start` - Start production server
- `npm run deploy` - Deploy to Cloudflare Workers
- `npm run lint` - Run ESLint
- `npm run format` - Format code with Prettier

## Integration with ZeroTwoApi

To integrate this Teams MCP server with your ZeroTwoApi backend:

### 1. Update Frontend OAuth Configuration

In your frontend OAuth providers configuration, add Teams:

```javascript
teams: {
  provider: "teams",
  clientId: import.meta.env.VITE_MICROSOFT_CLIENT_ID,
  flowType: OAuthFlowType.REDIRECT,
  scopes: [
    "Chat.Read",
    "ChannelMessage.Read.All",
    "User.Read",
  ],
  authEndpoint: "http://localhost:3003/oauth/authorize",
  backendExchangeEndpoint: "/api/auth/teams/callback",
  backendRefreshEndpoint: "/api/auth/teams/refresh",
  mcpEnabled: true,
  mcpEndpoint: "http://localhost:3003/mcp",
},
```

### 2. Update MCP Client Configuration

In your MCP client, add Teams MCP support:

```javascript
// For Teams MCP
if (integration.provider === "teams") {
  endpointUrl = "http://localhost:3003/mcp";
  
  // Use OAuth token from profile
  if (integration.oauth_token) {
    transportOptions.headers = {
      Authorization: `Bearer ${integration.oauth_token}`,
      ...integration.headers,
    };
  }
}
```

### 3. Create MCP Integration

After OAuth authentication, create an MCP integration in the database:

```sql
INSERT INTO mcp_integrations (
  user_id,
  name,
  provider,
  endpoint,
  mcp_type,
  connection_status,
  auth_type,
  description
) VALUES (
  'user_id_here',
  'Microsoft Teams',
  'teams',
  'http://localhost:3003/mcp',
  'remote_mcp',
  'connected',
  'oauth',
  'Microsoft Teams integration via MCP'
);
```

## Microsoft Graph API

This server uses Microsoft Graph API to interact with Teams:

- **Base URL**: `https://graph.microsoft.com/v1.0`
- **Documentation**: [Microsoft Graph API Reference](https://learn.microsoft.com/en-us/graph/api/overview)

### Common Graph API Endpoints Used

- `GET /me` - Get user profile
- `GET /me/chats` - List user's chats
- `GET /chats/{chat-id}/messages` - Get chat messages
- `GET /chats/{chat-id}/members` - Get chat members
- `GET /teams/{team-id}/channels/{channel-id}/messages` - Get channel messages

## Security Considerations

1. **OAuth Tokens**: Tokens are stored in memory (Node.js) or KV (Cloudflare). For production, ensure proper encryption.
2. **CORS**: Configure CORS origins appropriately for production.
3. **HTTPS**: Use HTTPS in production for secure communication.
4. **Environment Variables**: Never commit `.env` files. Use secure secret management in production.
5. **Rate Limiting**: Implement rate limiting for production deployments.
6. **Scope Minimization**: Only request the minimum scopes needed for your use case.

## Troubleshooting

### OAuth Errors

- **"Missing required environment variables"**: Check your `.env` file has all required variables
- **"Failed to exchange authorization code"**: Verify your redirect URI matches exactly in Azure Portal
- **"Token expired"**: The server automatically refreshes tokens if a refresh token is available
- **"Insufficient privileges"**: Ensure you've granted admin consent for the required API permissions

### MCP Connection Issues

- **"No valid session ID"**: Ensure the MCP client sends proper session headers
- **"Authentication required"**: User needs to complete OAuth flow first

### Graph API Errors

- **"Access denied"**: Check that the required permissions are granted and consented
- **"Resource not found"**: Verify the chat/channel/team IDs are correct
- **"Throttling"**: Microsoft Graph has rate limits. Implement backoff and retry logic.

## Rate Limits

Microsoft Graph API has rate limits:
- Per-app throttling: Varies by endpoint
- Per-user throttling: Varies by endpoint

See [Microsoft Graph throttling guidance](https://learn.microsoft.com/en-us/graph/throttling) for details.

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.








---

## Powered by ZeroTwo

This Microsoft Teams MCP connector is part of the [ZeroTwo AI platform](https://zerotwo.ai) — the all-in-one AI workspace that lets you send messages, manage channels, and automate your Teams collaboration through GPT-5, Claude, and Gemini.

| | |
|---|---|
| 🌐 **[ZeroTwo — All AI Models in One App](https://zerotwo.ai)** | Collaborate in Microsoft Teams with GPT-5, Claude, and Gemini — all in one workspace. |
| ✨ **[ZeroTwo Features](https://zerotwo.ai/features)** | AI messaging, document analysis, web search, and MCP-powered Teams collaboration tools. |
| 🤖 **[AI Models — GPT-5, Claude & Gemini](https://zerotwo.ai/zerotwo-models)** | Use the world's best AI to draft messages, summarize chats, and manage team channels. |
| 🔌 **[ZeroTwo Connectors & Integrations](https://zerotwo.ai/connectors)** | Connect Teams, Outlook, SharePoint, Gmail, and more to your AI workflow. |
| 💰 **[ZeroTwo Pricing](https://zerotwo.ai/pricing)** | One subscription that replaces ChatGPT Plus, Claude Pro, and Gemini Advanced. |
| 📝 **[ZeroTwo Blog](https://zerotwo.ai/blog)** | AI team collaboration tips, Microsoft 365 guides, and ZeroTwo product updates. |
| 🚀 **[Try ZeroTwo Free](https://app.zerotwo.ai/auth/login)** | Let AI supercharge your Teams workflow — get started free today. |

> **Built for ZeroTwo** — Use this Microsoft Teams MCP server with [ZeroTwo's AI connector system](https://zerotwo.ai/connectors) to send messages, manage channels, and automate your team communication through natural language.
