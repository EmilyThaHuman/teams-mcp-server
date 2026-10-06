/**
 * Microsoft Teams API Client wrapper
 */

import { logger } from '../utils/logger.js';
import {
  TeamsMessage,
  TeamsChat,
  TeamsChatMember,
  TeamsUser,
  SearchMessagesOptions,
  FetchMessageOptions,
} from '../types/index.js';

export class TeamsClient {
  private graphEndpoint = 'https://graph.microsoft.com/v1.0';

  /**
   * Make authenticated request to Microsoft Graph API
   */
  private async graphRequest(
    accessToken: string,
    endpoint: string,
    method: string = 'GET',
    body?: any
  ): Promise<any> {
    const url = endpoint.startsWith('http')
      ? endpoint
      : `${this.graphEndpoint}${endpoint}`;

    const options: RequestInit = {
      method,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    };

    if (body) {
      options.body = JSON.stringify(body);
    }

    const response = await fetch(url, options);

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph API request failed: ${response.status} - ${errorText}`);
    }

    return response.json();
  }

  /**
   * Search Teams messages
   */
  async searchMessages(
    accessToken: string,
    options: SearchMessagesOptions
  ): Promise<TeamsMessage[]> {
    try {
      const { query, chatId, channelId, teamId, maxResults = 20, fromDate, toDate } = options;

      logger.info('[TeamsClient] Searching messages', { query, chatId, channelId, teamId });

      const messages: TeamsMessage[] = [];

      // Search in specific chat
      if (chatId) {
        const endpoint = `/chats/${chatId}/messages`;
        const data = await this.graphRequest(accessToken, endpoint);
        const chatMessages = data.value || [];

        // Filter by query and date range
        const filtered = chatMessages.filter((msg: any) => {
          let matches = true;

          if (query) {
            const searchText = query.toLowerCase();
            const content = msg.body?.content?.toLowerCase() || '';
            const subject = msg.subject?.toLowerCase() || '';
            matches = content.includes(searchText) || subject.includes(searchText);
          }

          if (fromDate && msg.createdDateTime) {
            matches = matches && new Date(msg.createdDateTime) >= new Date(fromDate);
          }

          if (toDate && msg.createdDateTime) {
            matches = matches && new Date(msg.createdDateTime) <= new Date(toDate);
          }

          return matches;
        });

        messages.push(...filtered.slice(0, maxResults));
      }
      // Search in specific channel
      else if (teamId && channelId) {
        const endpoint = `/teams/${teamId}/channels/${channelId}/messages`;
        const data = await this.graphRequest(accessToken, endpoint);
        const channelMessages = data.value || [];

        // Filter by query and date range
        const filtered = channelMessages.filter((msg: any) => {
          let matches = true;

          if (query) {
            const searchText = query.toLowerCase();
            const content = msg.body?.content?.toLowerCase() || '';
            const subject = msg.subject?.toLowerCase() || '';
            matches = content.includes(searchText) || subject.includes(searchText);
          }

          if (fromDate && msg.createdDateTime) {
            matches = matches && new Date(msg.createdDateTime) >= new Date(fromDate);
          }

          if (toDate && msg.createdDateTime) {
            matches = matches && new Date(msg.createdDateTime) <= new Date(toDate);
          }

          return matches;
        });

        messages.push(...filtered.slice(0, maxResults));
      }
      // Search in all chats
      else {
        const chatsEndpoint = '/me/chats';
        const chatsData = await this.graphRequest(accessToken, chatsEndpoint);
        const chats = chatsData.value || [];

        for (const chat of chats.slice(0, 10)) {
          // Limit to first 10 chats
          try {
            const messagesEndpoint = `/chats/${chat.id}/messages`;
            const messagesData = await this.graphRequest(accessToken, messagesEndpoint);
            const chatMessages = messagesData.value || [];

            // Filter by query and date range
            const filtered = chatMessages.filter((msg: any) => {
              let matches = true;

              if (query) {
                const searchText = query.toLowerCase();
                const content = msg.body?.content?.toLowerCase() || '';
                const subject = msg.subject?.toLowerCase() || '';
                matches = content.includes(searchText) || subject.includes(searchText);
              }

              if (fromDate && msg.createdDateTime) {
                matches = matches && new Date(msg.createdDateTime) >= new Date(fromDate);
              }

              if (toDate && msg.createdDateTime) {
                matches = matches && new Date(msg.createdDateTime) <= new Date(toDate);
              }

              return matches;
            });

            messages.push(...filtered);

            if (messages.length >= maxResults) {
              break;
            }
          } catch (error) {
            logger.warn('[TeamsClient] Failed to fetch messages from chat', {
              chatId: chat.id,
              error,
            });
          }
        }
      }

      const result = messages.slice(0, maxResults);
      logger.info('[TeamsClient] Successfully searched messages', {
        count: result.length,
      });

      return result;
    } catch (error) {
      logger.error('[TeamsClient] Error searching messages:', error);
      throw new Error('Failed to search Teams messages');
    }
  }

  /**
   * Fetch a specific message by path
   */
  async fetchMessage(
    accessToken: string,
    options: FetchMessageOptions
  ): Promise<TeamsMessage> {
    try {
      const { path } = options;

      logger.info('[TeamsClient] Fetching message', { path });

      // Parse path to determine if it's a chat or channel message
      // Format: "chats/{chatId}/messages/{messageId}" or "teams/{teamId}/channels/{channelId}/messages/{messageId}"
      const endpoint = `/${path}`;
      const message = await this.graphRequest(accessToken, endpoint);

      logger.info('[TeamsClient] Successfully fetched message', { messageId: message.id });

      return message;
    } catch (error) {
      logger.error('[TeamsClient] Error fetching message:', error);
      throw new Error('Failed to fetch Teams message');
    }
  }

  /**
   * Get members of a chat
   */
  async getChatMembers(accessToken: string, chatId: string): Promise<TeamsChatMember[]> {
    try {
      logger.info('[TeamsClient] Getting chat members', { chatId });

      const endpoint = `/chats/${chatId}/members`;
      const data = await this.graphRequest(accessToken, endpoint);
      const members = data.value || [];

      logger.info('[TeamsClient] Successfully retrieved chat members', {
        count: members.length,
      });

      return members;
    } catch (error) {
      logger.error('[TeamsClient] Error getting chat members:', error);
      throw new Error('Failed to get chat members');
    }
  }

  /**
   * Get user profile
   */
  async getUserProfile(accessToken: string): Promise<TeamsUser> {
    try {
      logger.info('[TeamsClient] Getting user profile');

      const endpoint = '/me';
      const user = await this.graphRequest(accessToken, endpoint);

      logger.info('[TeamsClient] Successfully retrieved user profile', {
        userId: user.id,
        displayName: user.displayName,
      });

      return user;
    } catch (error) {
      logger.error('[TeamsClient] Error getting user profile:', error);
      throw new Error('Failed to get user profile');
    }
  }

  /**
   * Get user's chats
   */
  async getUserChats(accessToken: string, maxResults: number = 50): Promise<TeamsChat[]> {
    try {
      logger.info('[TeamsClient] Getting user chats');

      const endpoint = `/me/chats?$top=${maxResults}`;
      const data = await this.graphRequest(accessToken, endpoint);
      const chats = data.value || [];

      logger.info('[TeamsClient] Successfully retrieved user chats', {
        count: chats.length,
      });

      return chats;
    } catch (error) {
      logger.error('[TeamsClient] Error getting user chats:', error);
      throw new Error('Failed to get user chats');
    }
  }
}

// Singleton instance
export const teamsClient = new TeamsClient();







