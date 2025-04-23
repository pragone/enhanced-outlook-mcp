const winston = require('winston');
const config = require('../config');

// Create a custom format that includes timestamp, level, and message
const customFormat = winston.format.combine(
  winston.format.timestamp({
    format: 'YYYY-MM-DD HH:mm:ss'
  }),
  winston.format.printf(info => {
    return `${info.timestamp} [${info.level.toUpperCase()}]: ${info.message}${
      info.stack ? '\n' + info.stack : ''
    }${
      info.data ? '\n' + JSON.stringify(info.data, null, 2) : ''
    }`;
  })
);

// Create the logger
const logger = winston.createLogger({
  level: config.server.logLevel,
  format: customFormat,
  transports: [
    // Console transport removed to avoid interfering with STDIO JSON communication
    
    // File transport for errors and above
    new winston.transports.File({ 
      filename: 'error.log', 
      level: 'error' 
    }),
    // File transport for all logs
    new winston.transports.File({ 
      filename: 'combined.log' 
    })
  ]
});

// Add a method to log with attached data object
logger.logWithData = (level, message, data) => {
  logger.log(level, message, { data });
};

module.exports = logger;