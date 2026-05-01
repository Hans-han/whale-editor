import winston from 'winston';

const LOG_LEVEL = process.env.LOG_LEVEL ?? 'info';

export const logger = winston.createLogger({
  level: LOG_LEVEL,
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.json()
  ),
  transports: [
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      ),
    }),
  ],
});

export function truncateContent(content: string, maxLength = 200): string {
  if (content.length <= maxLength) return content;
  return content.slice(0, maxLength) + '...[truncated]';
}
