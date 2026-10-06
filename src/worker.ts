/**
 * Cloudflare Workers Entry Point for Microsoft Teams MCP Server
 * Implements MCP protocol with SSE support for ChatGPT integration
 */

import { teamsTools } from './mcp/tools.js';

// Environment interface for Cloudflare Workers
interface Env {
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_CLIENT_SECRET: string;
  MICROSOFT_REDIRECT_URI: string;
  MICROSOFT_TENANT_ID: string;
  FRONTEND_URL: string;
  BACKEND_API_URL?: string; // Optional: Backend API URL for token fallback
  MCP_SERVER_NAME: string;
  MCP_SERVER_VERSION: string;
  NODE_ENV: string;
  SESSIONS: KVNamespace;
}

// Session data structure
interface SessionData {
  userId: string;
  accessToken: string;
  refreshToken?: string;
  expiresAt?: string;
  createdAt: string;
}

// OAuth token structure
interface OAuthTokens {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  token_type?: string;
  scope?: string;
}

/**
 * Teams Client for Cloudflare Workers
 */
class TeamsClientWorker {
  private accessToken: string;
  private graphEndpoint = 'https://graph.microsoft.com/v1.0';

  constructor(accessToken: string) {
    this.accessToken = accessToken;
  }

  private async graphRequest(endpoint: string, method: string = 'GET'): Promise<any> {
    const url = endpoint.startsWith('http')
      ? endpoint
      : `${this.graphEndpoint}${endpoint}`;

    const response = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        'Content-Type': 'application/json',
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph API request failed: ${response.status} - ${errorText}`);
    }

    return response.json();
  }

  async searchMessages(options: any): Promise<any[]> {
    const { query, chatId, channelId, teamId, maxResults = 20 } = options;
    const messages: any[] = [];

    if (chatId) {
      const data = await this.graphRequest(`/chats/${chatId}/messages`);
      const chatMessages = data.value || [];
      const filtered = chatMessages.filter((msg: any) => {
        if (!query) return true;
        const searchText = query.toLowerCase();
        const content = msg.body?.content?.toLowerCase() || '';
        return content.includes(searchText);
      });
      messages.push(...filtered.slice(0, maxResults));
    } else if (teamId && channelId) {
      const data = await this.graphRequest(`/teams/${teamId}/channels/${channelId}/messages`);
      const channelMessages = data.value || [];
      const filtered = channelMessages.filter((msg: any) => {
        if (!query) return true;
        const searchText = query.toLowerCase();
        const content = msg.body?.content?.toLowerCase() || '';
        return content.includes(searchText);
      });
      messages.push(...filtered.slice(0, maxResults));
    } else {
      const chatsData = await this.graphRequest('/me/chats');
      const chats = chatsData.value || [];
      for (const chat of chats.slice(0, 10)) {
        try {
          const messagesData = await this.graphRequest(`/chats/${chat.id}/messages`);
          const chatMessages = messagesData.value || [];
          const filtered = chatMessages.filter((msg: any) => {
            if (!query) return true;
            const searchText = query.toLowerCase();
            const content = msg.body?.content?.toLowerCase() || '';
            return content.includes(searchText);
          });
          messages.push(...filtered);
          if (messages.length >= maxResults) break;
        } catch (error) {
          console.warn('Failed to fetch messages from chat:', chat.id);
        }
      }
    }

    return messages.slice(0, maxResults);
  }

  async fetchMessage(path: string): Promise<any> {
    return this.graphRequest(`/${path}`);
  }

  async getChatMembers(chatId: string): Promise<any[]> {
    const data = await this.graphRequest(`/chats/${chatId}/members`);
    return data.value || [];
  }

  async getUserProfile(): Promise<any> {
    return this.graphRequest('/me');
  }
}

/**
 * OAuth Manager for Cloudflare Workers
 */
class CloudflareOAuthManager {
  private env: Env;
  private authority: string;
  private tokenEndpoint: string;
  private authEndpoint: string;

  constructor(env: Env) {
    this.env = env;
    this.authority = `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}`;
    this.tokenEndpoint = `${this.authority}/oauth2/v2.0/token`;
    this.authEndpoint = `${this.authority}/oauth2/v2.0/authorize`;
  }

  getAuthorizationUrl(state: string): string {
    const params = new URLSearchParams({
      client_id: this.env.MICROSOFT_CLIENT_ID,
      response_type: 'code',
      redirect_uri: this.env.MICROSOFT_REDIRECT_URI,
      response_mode: 'query',
      scope: 'Chat.Read ChannelMessage.Read.All User.Read',
      state,
      prompt: 'consent',
    });
    return `${this.authEndpoint}?${params.toString()}`;
  }

  async exchangeCodeForTokens(code: string): Promise<OAuthTokens> {
    const params = new URLSearchParams({
      client_id: this.env.MICROSOFT_CLIENT_ID,
      client_secret: this.env.MICROSOFT_CLIENT_SECRET,
      code,
      redirect_uri: this.env.MICROSOFT_REDIRECT_URI,
      grant_type: 'authorization_code',
      scope: 'Chat.Read ChannelMessage.Read.All User.Read',
    });

    const response = await fetch(this.tokenEndpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });

    if (!response.ok) {
      throw new Error(`Failed to exchange code: ${await response.text()}`);
    }

    const data = await response.json();
    return {
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      expires_in: data.expires_in || 3600,
      token_type: data.token_type || 'Bearer',
      scope: data.scope,
    };
  }

  async refreshAccessToken(refreshToken: string): Promise<OAuthTokens> {
    const params = new URLSearchParams({
      client_id: this.env.MICROSOFT_CLIENT_ID,
      client_secret: this.env.MICROSOFT_CLIENT_SECRET,
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
      scope: 'Chat.Read ChannelMessage.Read.All User.Read',
    });

    const response = await fetch(this.tokenEndpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });

    if (!response.ok) {
      throw new Error(`Failed to refresh token: ${await response.text()}`);
    }

    const data = await response.json();
    return {
      access_token: data.access_token,
      refresh_token: data.refresh_token || refreshToken,
      expires_in: data.expires_in || 3600,
      token_type: data.token_type || 'Bearer',
      scope: data.scope,
    };
  }

  async storeSession(userId: string, tokens: OAuthTokens): Promise<void> {
    const expiresAt = tokens.expires_in
      ? new Date(Date.now() + tokens.expires_in * 1000).toISOString()
      : undefined;

    const sessionData: SessionData = {
      userId,
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresAt,
      createdAt: new Date().toISOString(),
    };

    await this.env.SESSIONS.put(
      `session:${userId}`,
      JSON.stringify(sessionData),
      { expirationTtl: 60 * 60 * 24 * 30 }
    );
  }

  async getSession(userId: string): Promise<SessionData | null> {
    const data = await this.env.SESSIONS.get(`session:${userId}`);
    return data ? JSON.parse(data) : null;
  }

  async isSessionValid(userId: string): Promise<boolean> {
    const session = await this.getSession(userId);
    if (!session) return false;
    if (!session.expiresAt) return true;
    return new Date(session.expiresAt) > new Date();
  }

  async getValidAccessToken(userId: string): Promise<string | null> {
    const session = await this.getSession(userId);
    if (session && await this.isSessionValid(userId)) {
      return session.accessToken;
    }

    if (session && session.refreshToken) {
      try {
        const newTokens = await this.refreshAccessToken(session.refreshToken);
        await this.storeSession(userId, newTokens);
        return newTokens.access_token;
      } catch (error) {
        console.error('Failed to refresh token from KV:', error);
      }
    }

    try {
      const backendUrl = this.env.BACKEND_API_URL || 
                        (this.env.FRONTEND_URL ? this.env.FRONTEND_URL.replace(/\/$/, '') : null) ||
                        'https://api.zerotwo.app';
      
      const response = await fetch(`${backendUrl}/api/ai/tools/teams/token?userId=${userId}`, {
        method: 'GET',
        headers: { 'Content-Type': 'application/json' },
      });

      if (response.ok) {
        const data = await response.json() as {
          accessToken?: string;
          refreshToken?: string;
          expiresIn?: number;
        };
        if (data.accessToken) {
          await this.storeSession(userId, {
            access_token: data.accessToken,
            refresh_token: data.refreshToken,
            expires_in: data.expiresIn,
            token_type: 'Bearer',
          });
          return data.accessToken;
        }
      }
    } catch (error) {
      console.error('Failed to fetch token from backend API (fallback):', error);
    }

    return null;
  }

  async removeSession(userId: string): Promise<void> {
    await this.env.SESSIONS.delete(`session:${userId}`);
  }
}

/**
 * CORS headers
 */
function getCorsHeaders(origin?: string): Record<string, string> {
  const allowedOrigins = [
    'http://localhost:5173',
    'http://localhost:3000',
    'https://zerotwo.app',
  ];

  const requestOrigin = origin || '';
  const allowOrigin = allowedOrigins.includes(requestOrigin)
    ? requestOrigin
    : allowedOrigins[0];

  return {
    'Access-Control-Allow-Origin': allowOrigin,
    'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, mcp-session-id, Authorization',
    'Access-Control-Expose-Headers': 'Mcp-Session-Id',
    'Access-Control-Allow-Credentials': 'true',
  };
}

/**
 * Handle OPTIONS requests
 */
function handleOptions(request: Request): Response {
  return new Response(null, {
    status: 204,
    headers: getCorsHeaders(request.headers.get('Origin') || undefined),
  });
}

/**
 * Execute MCP tool
 */
async function executeTool(
  toolName: string,
  args: any,
  oauthManager: CloudflareOAuthManager,
  env: Env,
  requestHeaders?: Headers
): Promise<any> {
  let userId = args?.userId;
  if (!userId && requestHeaders) {
    userId = requestHeaders.get('X-User-Id');
  }
  
  if (!userId) {
    return {
      content: [
        {
          type: 'text',
          text: 'User ID is required for all Teams operations. Please provide userId in tool arguments or X-User-Id header.',
        },
      ],
      isError: true,
    };
  }

  // First, try to get access token from X-Access-Token header (passed directly from backend)
  let accessToken = requestHeaders?.get('X-Access-Token') || null;
  
  // If not in header, fall back to KV storage lookup
  if (!accessToken) {
    accessToken = await oauthManager.getValidAccessToken(userId);
  }
  
  if (!accessToken) {
    return {
      content: [
        {
          type: 'text',
          text: 'Authentication required. Please authenticate with Microsoft Teams first.',
        },
      ],
      isError: true,
    };
  }

  const teamsClient = new TeamsClientWorker(accessToken);

  try {
    switch (toolName) {
      case 'teams_search': {
        const { query, chatId, channelId, teamId, maxResults, fromDate, toDate } = args;
        const messages = await teamsClient.searchMessages({
          query,
          chatId,
          channelId,
          teamId,
          maxResults,
          fromDate,
          toDate,
        });

        const output = {
          messages: messages.map((msg: any) => ({
            id: msg.id,
            chatId: msg.chatId,
            channelId: msg.channelId,
            teamId: msg.teamId,
            from: {
              displayName: msg.from?.user?.displayName || msg.from?.application?.displayName || 'Unknown',
              id: msg.from?.user?.id || msg.from?.application?.id,
            },
            body: msg.body,
            createdDateTime: msg.createdDateTime,
            subject: msg.subject,
            webUrl: msg.webUrl,
          })),
          count: messages.length,
        };

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      case 'teams_fetch': {
        const { path } = args;
        const message = await teamsClient.fetchMessage(path);

        const output = {
          id: message.id,
          chatId: message.chatId,
          channelId: message.channelId,
          teamId: message.teamId,
          from: {
            displayName: message.from?.user?.displayName || message.from?.application?.displayName || 'Unknown',
            id: message.from?.user?.id || message.from?.application?.id,
          },
          body: message.body,
          createdDateTime: message.createdDateTime,
          subject: message.subject,
          webUrl: message.webUrl,
        };

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      case 'teams_get_chat_members': {
        const { chatId } = args;
        const members = await teamsClient.getChatMembers(chatId);

        const output = {
          members: members.map((member: any) => ({
            id: member.id,
            displayName: member.displayName,
            userId: member.userId,
            email: member.email,
            roles: member.roles,
          })),
          count: members.length,
        };

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      case 'teams_get_profile': {
        const profile = await teamsClient.getUserProfile();

        const output = {
          id: profile.id,
          displayName: profile.displayName,
          givenName: profile.givenName,
          surname: profile.surname,
          userPrincipalName: profile.userPrincipalName,
          mail: profile.mail,
          jobTitle: profile.jobTitle,
          officeLocation: profile.officeLocation,
          mobilePhone: profile.mobilePhone,
          businessPhones: profile.businessPhones,
        };

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      default:
        return {
          content: [
            {
              type: 'text',
              text: `Unknown tool: ${toolName}`,
            },
          ],
          isError: true,
        };
    }
  } catch (error: any) {
    console.error(`Error executing tool ${toolName}:`, error);
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${error.message}`,
        },
      ],
      isError: true,
    };
  }
}

/**
 * Handle MCP protocol requests
 */
async function handleMcpRequest(
  request: Request,
  env: Env,
  oauthManager: CloudflareOAuthManager
): Promise<Response> {
  const requestHeaders = request.headers;
  try {
    const body = await request.json() as {
      method?: string;
      id?: string | number;
      params?: {
        name?: string;
        arguments?: any;
      };
    };

    if (body.method === 'initialize') {
      return new Response(
        JSON.stringify({
          jsonrpc: '2.0',
          id: body.id,
          result: {
            protocolVersion: '2024-11-05',
            capabilities: {
              tools: {},
            },
            serverInfo: {
              name: env.MCP_SERVER_NAME || 'teams-mcp-server',
              version: env.MCP_SERVER_VERSION || '1.0.0',
            },
          },
        }),
        {
          headers: {
            'Content-Type': 'application/json',
            ...getCorsHeaders(request.headers.get('Origin') || undefined),
          },
        }
      );
    }

    if (body.method === 'tools/list') {
      const tools = teamsTools.map((tool) => {
        const schema = tool.definition.inputSchema as any;
        const properties: Record<string, any> = {};
        
        if (schema && typeof schema === 'object') {
          if (schema.shape) {
            Object.entries(schema.shape).forEach(([key, value]: [string, any]) => {
              properties[key] = {
                type: value._def?.typeName === 'ZodString' ? 'string' : 
                      value._def?.typeName === 'ZodNumber' ? 'number' :
                      value._def?.typeName === 'ZodBoolean' ? 'boolean' :
                      value._def?.typeName === 'ZodArray' ? 'array' : 'string',
                description: value.description || value._def?.description || '',
              };
            });
          }
        }

        return {
          name: tool.name,
          description: tool.definition.description || tool.definition.title || '',
          inputSchema: {
            type: 'object',
            properties,
            required: [],
          },
        };
      });

      return new Response(
        JSON.stringify({
          jsonrpc: '2.0',
          id: body.id,
          result: {
            tools,
          },
        }),
        {
          headers: {
            'Content-Type': 'application/json',
            ...getCorsHeaders(request.headers.get('Origin') || undefined),
          },
        }
      );
    }

    if (body.method === 'tools/call') {
      if (!body.params) {
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            id: body.id,
            error: {
              code: -32602,
              message: 'Invalid params',
            },
          }),
          {
            status: 400,
            headers: {
              'Content-Type': 'application/json',
              ...getCorsHeaders(),
            },
          }
        );
      }

      const { name, arguments: args } = body.params;
      if (!name) {
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            id: body.id,
            error: {
              code: -32602,
              message: 'Tool name is required',
            },
          }),
          {
            status: 400,
            headers: {
              'Content-Type': 'application/json',
              ...getCorsHeaders(),
            },
          }
        );
      }

      const result = await executeTool(name, args || {}, oauthManager, env, request.headers);

      return new Response(
        JSON.stringify({
          jsonrpc: '2.0',
          id: body.id,
          result: {
            content: result.content,
            isError: result.isError || false,
          },
        }),
        {
          headers: {
            'Content-Type': 'application/json',
            ...getCorsHeaders(request.headers.get('Origin') || undefined),
          },
        }
      );
    }

    // Handle notifications (no id field, no response expected)
    // Notifications like notifications/initialized don't have an id and don't require a JSON-RPC response
    // Use 204 No Content for proper HTTP semantics
    if (body.method?.startsWith('notifications/') || (!body.id && body.method)) {
      // For notifications, just acknowledge with 204 No Content
      return new Response(null, {
        status: 204,
        headers: getCorsHeaders(request.headers.get('Origin') || undefined),
      });
    }

    // Unknown method (only for requests with id)
    return new Response(
      JSON.stringify({
        jsonrpc: '2.0',
        id: body.id,
        error: {
          code: -32601,
          message: 'Method not found',
        },
      }),
      {
        status: 400,
        headers: {
          'Content-Type': 'application/json',
          ...getCorsHeaders(request.headers.get('Origin') || undefined),
        },
      }
    );
  } catch (error: any) {
    console.error('MCP request error:', error);
    return new Response(
      JSON.stringify({
        jsonrpc: '2.0',
        id: null,
        error: {
          code: -32603,
          message: 'Internal error',
          data: error.message,
        },
      }),
      {
        status: 500,
        headers: {
          'Content-Type': 'application/json',
          ...getCorsHeaders(request.headers.get('Origin') || undefined),
        },
      }
    );
  }
}

/**
 * SSE endpoint for MCP protocol
 */
function createSseResponse(env: Env): Response {
  const encoder = new TextEncoder();
  let keepAliveInterval: ReturnType<typeof setInterval> | null = null;

  const stream = new ReadableStream({
    start(controller) {
      const initMessage = {
        jsonrpc: '2.0',
        method: 'notifications/initialized',
        params: {},
      };
      controller.enqueue(encoder.encode(`data: ${JSON.stringify(initMessage)}\n\n`));

      keepAliveInterval = setInterval(() => {
        try {
          controller.enqueue(encoder.encode(': keepalive\n\n'));
        } catch (e) {
          if (keepAliveInterval) {
            clearInterval(keepAliveInterval);
            keepAliveInterval = null;
          }
        }
      }, 30000);
    },
    cancel() {
      if (keepAliveInterval) {
        clearInterval(keepAliveInterval);
        keepAliveInterval = null;
      }
    },
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      Connection: 'keep-alive',
      ...getCorsHeaders(),
    },
  });
}

/**
 * Main Worker fetch handler
 */
export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin');
    const corsHeaders = getCorsHeaders(origin || undefined);

    if (request.method === 'OPTIONS') {
      return handleOptions(request);
    }

    try {
      const oauthManager = new CloudflareOAuthManager(env);

      if (url.pathname === '/health' || url.pathname === '/') {
        return new Response(
          JSON.stringify({
            status: 'healthy',
            server: env.MCP_SERVER_NAME,
            version: env.MCP_SERVER_VERSION,
            timestamp: new Date().toISOString(),
          }),
          {
            headers: {
              'Content-Type': 'application/json',
              ...corsHeaders,
            },
          }
        );
      }

      if (url.pathname === '/oauth/authorize' && request.method === 'GET') {
        const userId = url.searchParams.get('userId');
        const state = url.searchParams.get('state');

        if (!userId || !state) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameters: userId and state' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        const authUrl = oauthManager.getAuthorizationUrl(state);

        return new Response(
          JSON.stringify({ authorizationUrl: authUrl, state }),
          {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
          }
        );
      }

      if (url.pathname === '/oauth/callback' && request.method === 'GET') {
        const code = url.searchParams.get('code');
        const state = url.searchParams.get('state');
        const error = url.searchParams.get('error');

        if (error) {
          return Response.redirect(
            `${env.FRONTEND_URL}/settings?oauth_error=${encodeURIComponent(error)}`,
            302
          );
        }

        if (!code || !state) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameters: code and state' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        const tokens = await oauthManager.exchangeCodeForTokens(code);
        const userId = state;

        await oauthManager.storeSession(userId, tokens);

        return Response.redirect(
          `${env.FRONTEND_URL}/settings?oauth_success=true&provider=teams`,
          302
        );
      }

      if (url.pathname === '/oauth/refresh' && request.method === 'POST') {
        const body = await request.json() as { userId?: string; refreshToken?: string };
        const { userId, refreshToken } = body;

        if (!userId || !refreshToken) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameters: userId and refreshToken' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        const tokens = await oauthManager.refreshAccessToken(refreshToken);
        await oauthManager.storeSession(userId, tokens);

        return new Response(
          JSON.stringify({
            accessToken: tokens.access_token,
            refreshToken: tokens.refresh_token,
            expiresIn: tokens.expires_in,
            timestamp: new Date().toISOString(),
          }),
          {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
          }
        );
      }

      if (url.pathname === '/oauth/disconnect' && request.method === 'POST') {
        const body = await request.json() as { userId?: string };
        const { userId } = body;

        if (!userId) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameter: userId' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        await oauthManager.removeSession(userId);

        return new Response(
          JSON.stringify({ success: true, message: 'Successfully disconnected' }),
          {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
          }
        );
      }

      if (url.pathname === '/oauth/sync' && request.method === 'POST') {
        const body = await request.json() as {
          userId?: string;
          accessToken?: string;
          refreshToken?: string;
          expiresIn?: number;
        };
        const { userId, accessToken, refreshToken, expiresIn } = body;

        if (!userId || !accessToken) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameters: userId and accessToken' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        await oauthManager.storeSession(userId, {
          access_token: accessToken,
          refresh_token: refreshToken,
          expires_in: expiresIn,
          token_type: 'Bearer',
        });

        return new Response(
          JSON.stringify({ success: true, message: 'Tokens synced successfully' }),
          {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
          }
        );
      }

      if (url.pathname === '/mcp' && request.method === 'POST') {
        return handleMcpRequest(request, env, oauthManager);
      }

      if (url.pathname === '/sse' && request.method === 'GET') {
        return createSseResponse(env);
      }

      return new Response(
        JSON.stringify({ error: 'Not Found', message: 'The requested endpoint does not exist' }),
        {
          status: 404,
          headers: { 'Content-Type': 'application/json', ...corsHeaders },
        }
      );
    } catch (error: any) {
      console.error('Worker error:', error);
      return new Response(
        JSON.stringify({
          error: 'Internal Server Error',
          message: error.message,
        }),
        {
          status: 500,
          headers: { 'Content-Type': 'application/json', ...corsHeaders },
        }
      );
    }
  },
};







