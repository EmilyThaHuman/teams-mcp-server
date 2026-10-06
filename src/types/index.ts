/**
 * Type definitions for Microsoft Teams MCP Server
 */

export interface TeamsMessage {
  id: string;
  chatId?: string;
  channelId?: string;
  teamId?: string;
  from: {
    user?: {
      id: string;
      displayName: string;
      userPrincipalName?: string;
    };
    application?: {
      id: string;
      displayName: string;
    };
  };
  body: {
    content: string;
    contentType: 'text' | 'html';
  };
  createdDateTime: string;
  lastModifiedDateTime?: string;
  deletedDateTime?: string;
  subject?: string;
  summary?: string;
  importance: 'normal' | 'high' | 'urgent';
  messageType: 'message' | 'chatEvent' | 'typing' | 'unknownFutureValue';
  webUrl?: string;
  attachments?: TeamsAttachment[];
  mentions?: TeamsMention[];
}

export interface TeamsAttachment {
  id: string;
  contentType: string;
  contentUrl?: string;
  content?: string;
  name?: string;
  thumbnailUrl?: string;
}

export interface TeamsMention {
  id: number;
  mentionText: string;
  mentioned: {
    user?: {
      id: string;
      displayName: string;
      userPrincipalName?: string;
    };
  };
}

export interface TeamsChat {
  id: string;
  topic?: string;
  createdDateTime: string;
  lastUpdatedDateTime: string;
  chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknownFutureValue';
  webUrl?: string;
  tenantId?: string;
  members?: TeamsChatMember[];
}

export interface TeamsChatMember {
  id: string;
  displayName: string;
  userId?: string;
  email?: string;
  roles?: string[];
  visibleHistoryStartDateTime?: string;
}

export interface TeamsChannel {
  id: string;
  displayName: string;
  description?: string;
  email?: string;
  webUrl?: string;
  membershipType?: 'standard' | 'private' | 'unknownFutureValue';
}

export interface TeamsTeam {
  id: string;
  displayName: string;
  description?: string;
  internalId?: string;
  webUrl?: string;
}

export interface TeamsUser {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  userPrincipalName?: string;
  mail?: string;
  jobTitle?: string;
  officeLocation?: string;
  mobilePhone?: string;
  businessPhones?: string[];
}

export interface OAuthTokens {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  token_type?: string;
  scope?: string;
  id_token?: string;
}

export interface SessionData {
  userId: string;
  accessToken: string;
  refreshToken?: string;
  expiresAt?: Date;
  createdAt: Date;
}

export interface SearchMessagesOptions {
  query: string;
  chatId?: string;
  channelId?: string;
  teamId?: string;
  maxResults?: number;
  fromDate?: string;
  toDate?: string;
}

export interface FetchMessageOptions {
  path: string; // Format: "chats/{chatId}/messages/{messageId}" or "teams/{teamId}/channels/{channelId}/messages/{messageId}"
}







