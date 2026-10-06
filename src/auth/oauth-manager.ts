/**
 * OAuth Manager for handling Microsoft OAuth 2.0 authentication
 */

import { config } from '../config/index.js';
import { logger } from '../utils/logger.js';
import { OAuthTokens, SessionData } from '../types/index.js';

export class OAuthManager {
  private sessions: Map<string, SessionData>;
  private authority: string;
  private tokenEndpoint: string;
  private authEndpoint: string;

  constructor() {
    this.sessions = new Map();
    this.authority = `https://login.microsoftonline.com/${config.microsoft.tenantId}`;
    this.tokenEndpoint = `${this.authority}/oauth2/v2.0/token`;
    this.authEndpoint = `${this.authority}/oauth2/v2.0/authorize`;
  }

  /**
   * Generate OAuth authorization URL
   */
  getAuthorizationUrl(state: string): string {
    const params = new URLSearchParams({
      client_id: config.microsoft.clientId,
      response_type: 'code',
      redirect_uri: config.microsoft.redirectUri,
      response_mode: 'query',
      scope: config.teams.scopes.join(' '),
      state,
      prompt: 'consent', // Force consent to get refresh token
    });

    const authUrl = `${this.authEndpoint}?${params.toString()}`;
    logger.info('[OAuthManager] Generated authorization URL', { state });
    return authUrl;
  }

  /**
   * Exchange authorization code for tokens
   */
  async exchangeCodeForTokens(code: string): Promise<OAuthTokens> {
    try {
      logger.info('[OAuthManager] Exchanging code for tokens');

      const params = new URLSearchParams({
        client_id: config.microsoft.clientId,
        client_secret: config.microsoft.clientSecret,
        code,
        redirect_uri: config.microsoft.redirectUri,
        grant_type: 'authorization_code',
        scope: config.teams.scopes.join(' '),
      });

      const response = await fetch(this.tokenEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString(),
      });

      if (!response.ok) {
        const error = await response.text();
        throw new Error(`Failed to exchange code: ${error}`);
      }

      const data = await response.json();

      logger.info('[OAuthManager] Successfully exchanged code for tokens');

      return {
        access_token: data.access_token,
        refresh_token: data.refresh_token,
        expires_in: data.expires_in || 3600,
        token_type: data.token_type || 'Bearer',
        scope: data.scope,
        id_token: data.id_token,
      };
    } catch (error) {
      logger.error('[OAuthManager] Error exchanging code for tokens:', error);
      throw new Error('Failed to exchange authorization code for tokens');
    }
  }

  /**
   * Refresh access token using refresh token
   */
  async refreshAccessToken(refreshToken: string): Promise<OAuthTokens> {
    try {
      logger.info('[OAuthManager] Refreshing access token');

      const params = new URLSearchParams({
        client_id: config.microsoft.clientId,
        client_secret: config.microsoft.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: config.teams.scopes.join(' '),
      });

      const response = await fetch(this.tokenEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString(),
      });

      if (!response.ok) {
        const error = await response.text();
        throw new Error(`Failed to refresh token: ${error}`);
      }

      const data = await response.json();

      logger.info('[OAuthManager] Successfully refreshed access token');

      return {
        access_token: data.access_token,
        refresh_token: data.refresh_token || refreshToken, // Keep the same refresh token if not provided
        expires_in: data.expires_in || 3600,
        token_type: data.token_type || 'Bearer',
        scope: data.scope,
        id_token: data.id_token,
      };
    } catch (error) {
      logger.error('[OAuthManager] Error refreshing access token:', error);
      throw new Error('Failed to refresh access token');
    }
  }

  /**
   * Store session data
   */
  storeSession(userId: string, tokens: OAuthTokens): void {
    const expiresAt = tokens.expires_in
      ? new Date(Date.now() + tokens.expires_in * 1000)
      : undefined;

    const sessionData: SessionData = {
      userId,
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresAt,
      createdAt: new Date(),
    };

    this.sessions.set(userId, sessionData);
    logger.info('[OAuthManager] Stored session for user', { userId });
  }

  /**
   * Get session data
   */
  getSession(userId: string): SessionData | undefined {
    return this.sessions.get(userId);
  }

  /**
   * Check if session is valid (not expired)
   */
  isSessionValid(userId: string): boolean {
    const session = this.sessions.get(userId);
    if (!session) {
      return false;
    }

    if (!session.expiresAt) {
      return true; // No expiration set
    }

    return session.expiresAt > new Date();
  }

  /**
   * Get valid access token (refresh if needed)
   */
  async getValidAccessToken(userId: string): Promise<string | null> {
    const session = this.sessions.get(userId);
    if (!session) {
      logger.warn('[OAuthManager] No session found for user', { userId });
      return null;
    }

    // Check if token is still valid
    if (this.isSessionValid(userId)) {
      return session.accessToken;
    }

    // Token expired, try to refresh
    if (!session.refreshToken) {
      logger.warn('[OAuthManager] No refresh token available for user', { userId });
      return null;
    }

    try {
      const newTokens = await this.refreshAccessToken(session.refreshToken);
      this.storeSession(userId, newTokens);
      return newTokens.access_token;
    } catch (error) {
      logger.error('[OAuthManager] Failed to refresh token for user', { userId, error });
      return null;
    }
  }

  /**
   * Remove session
   */
  removeSession(userId: string): void {
    this.sessions.delete(userId);
    logger.info('[OAuthManager] Removed session for user', { userId });
  }

  /**
   * Get all active sessions
   */
  getActiveSessions(): string[] {
    return Array.from(this.sessions.keys());
  }
}

// Singleton instance
export const oauthManager = new OAuthManager();







