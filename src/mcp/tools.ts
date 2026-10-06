/**
 * MCP Tool definitions for Microsoft Teams operations
 */

import { z } from 'zod';
import { teamsClient } from '../teams/client.js';
import { oauthManager } from '../auth/oauth-manager.js';
import { logger } from '../utils/logger.js';

/**
 * Teams Search Messages Tool
 */
export const teamsSearchTool = {
  name: 'teams_search',
  definition: {
    title: 'Search Microsoft Teams chats and channel messages',
    description: 'Use this to find Microsoft Teams messages across chats or channels. It returns message IDs plus chat, channel, and team identifiers that the model can use to build the `path` required by `teams_fetch`.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      query: z.string().describe('Search query text'),
      chatId: z.string().optional().describe('Specific chat ID to search in'),
      channelId: z.string().optional().describe('Specific channel ID to search in'),
      teamId: z.string().optional().describe('Team ID (required if channelId is provided)'),
      maxResults: z.number().min(1).max(100).default(20).describe('Maximum number of messages to return'),
      fromDate: z.string().optional().describe('Filter messages from this date (ISO 8601 format)'),
      toDate: z.string().optional().describe('Filter messages to this date (ISO 8601 format)'),
    },
    outputSchema: {
      messages: z.array(
        z.object({
          id: z.string(),
          chatId: z.string().optional(),
          channelId: z.string().optional(),
          teamId: z.string().optional(),
          from: z.object({
            displayName: z.string(),
            id: z.string().optional(),
          }),
          body: z.object({
            content: z.string(),
            contentType: z.string(),
          }),
          createdDateTime: z.string(),
          subject: z.string().optional(),
          webUrl: z.string().optional(),
        })
      ),
      count: z.number(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, query, chatId, channelId, teamId, maxResults, fromDate, toDate } = args;

      logger.info('[Tool:teams_search] Executing', { userId, query, chatId });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // Search messages
      const messages = await teamsClient.searchMessages(accessToken, {
        query,
        chatId,
        channelId,
        teamId,
        maxResults,
        fromDate,
        toDate,
      });

      const output = {
        messages: messages.map((msg) => ({
          id: msg.id,
          chatId: msg.chatId,
          channelId: msg.channelId,
          teamId: msg.teamId,
          from: {
            displayName: msg.from.user?.displayName || msg.from.application?.displayName || 'Unknown',
            id: msg.from.user?.id || msg.from.application?.id,
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
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:teams_search] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error searching messages: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Teams Fetch Message Tool
 */
export const teamsFetchTool = {
  name: 'teams_fetch',
  definition: {
    title: 'Fetch a Teams message by path',
    description: 'Use this to read one exact Teams message when you already know its message path. Build the `path` from `teams_search` results as `chats/{chatId}/messages/{messageId}` for chats or `teams/{teamId}/channels/{channelId}/messages/{messageId}` for channel messages.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      path: z
        .string()
        .describe(
          'Message path to the exact Teams message. Build this from `teams_search` results using `chats/{chatId}/messages/{messageId}` or `teams/{teamId}/channels/{channelId}/messages/{messageId}`.'
        ),
    },
    outputSchema: {
      id: z.string(),
      chatId: z.string().optional(),
      channelId: z.string().optional(),
      teamId: z.string().optional(),
      from: z.object({
        displayName: z.string(),
        id: z.string().optional(),
      }),
      body: z.object({
        content: z.string(),
        contentType: z.string(),
      }),
      createdDateTime: z.string(),
      subject: z.string().optional(),
      webUrl: z.string().optional(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, path } = args;

      logger.info('[Tool:teams_fetch] Executing', { userId, path });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // Fetch message
      const message = await teamsClient.fetchMessage(accessToken, { path });

      const output = {
        id: message.id,
        chatId: message.chatId,
        channelId: message.channelId,
        teamId: message.teamId,
        from: {
          displayName: message.from.user?.displayName || message.from.application?.displayName || 'Unknown',
          id: message.from.user?.id || message.from.application?.id,
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
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:teams_fetch] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error fetching message: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Teams Get Chat Members Tool
 */
export const teamsGetChatMembersTool = {
  name: 'teams_get_chat_members',
  definition: {
    title: 'List the members of a Teams chat',
    description: 'Use this when you already know a Teams `chatId` and need the participant list for that chat, for example after locating the chat from message results.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      chatId: z.string().describe('Teams chat ID, typically taken from a `teams_search` result.'),
    },
    outputSchema: {
      members: z.array(
        z.object({
          id: z.string(),
          displayName: z.string(),
          userId: z.string().optional(),
          email: z.string().optional(),
          roles: z.array(z.string()).optional(),
        })
      ),
      count: z.number(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, chatId } = args;

      logger.info('[Tool:teams_get_chat_members] Executing', { userId, chatId });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // Get chat members
      const members = await teamsClient.getChatMembers(accessToken, chatId);

      const output = {
        members: members.map((member) => ({
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
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:teams_get_chat_members] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error getting chat members: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Teams Get Profile Tool
 */
export const teamsGetProfileTool = {
  name: 'teams_get_profile',
  definition: {
    title: "Return the authenticated Teams user's profile",
    description: 'Use this when you need the authenticated Microsoft Teams user profile and account metadata. It does not search chats or messages.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
    },
    outputSchema: {
      id: z.string(),
      displayName: z.string(),
      givenName: z.string().optional(),
      surname: z.string().optional(),
      userPrincipalName: z.string().optional(),
      mail: z.string().optional(),
      jobTitle: z.string().optional(),
      officeLocation: z.string().optional(),
      mobilePhone: z.string().optional(),
      businessPhones: z.array(z.string()).optional(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId } = args;

      logger.info('[Tool:teams_get_profile] Executing', { userId });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // Get user profile
      const profile = await teamsClient.getUserProfile(accessToken);

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
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:teams_get_profile] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error getting user profile: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Export all tools
 */
export const teamsTools = [
  teamsSearchTool,
  teamsFetchTool,
  teamsGetChatMembersTool,
  teamsGetProfileTool,
];






