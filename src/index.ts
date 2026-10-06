/**
 * Microsoft Teams MCP Server - Entry point
 */

import { config, validateConfig } from './config/index.js';
import { logger } from './utils/logger.js';
import { TeamsMcpServer } from './mcp/server.js';

async function main() {
  try {
    logger.info('Starting Microsoft Teams MCP Server...');

    // Validate configuration
    validateConfig();
    logger.info('Configuration validated');

    // Create and start server
    const server = new TeamsMcpServer();
    server.start();

    // Graceful shutdown
    process.on('SIGTERM', () => {
      logger.info('SIGTERM received, shutting down gracefully...');
      process.exit(0);
    });

    process.on('SIGINT', () => {
      logger.info('SIGINT received, shutting down gracefully...');
      process.exit(0);
    });
  } catch (error: any) {
    logger.error('Failed to start server:', error);
    process.exit(1);
  }
}

main();







